using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging.Archive;

/// <summary>Loads OFD containers with one cumulative bounded-read accounting context.</summary>
public sealed partial class OfdPackageLoader
{
    internal static OfdPackageLoadOptions SnapshotOptions(OfdPackageLoadOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        var snapshot = new OfdPackageLoadOptions
        {
            MaxInputBytes = options.MaxInputBytes, MaxEntryCount = options.MaxEntryCount, MaxPageCount = options.MaxPageCount,
            MaxAnnotationObjectCount = options.MaxAnnotationObjectCount, MaxAnnotationXmlBytes = options.MaxAnnotationXmlBytes,
            MaxEntryUncompressedBytes = options.MaxEntryUncompressedBytes, MaxTotalUncompressedBytes = options.MaxTotalUncompressedBytes,
            MaxCompressionRatio = options.MaxCompressionRatio
        };
        ValidateOptions(snapshot); return snapshot;
    }

    // This controlled profile is reviewed for .NET 10 Read-mode ZipArchive.
    // It screens a fixed footer only; ZipArchive remains the ZIP parser.
    internal async Task<OfdPackageArchive> LoadAccountedAsync(Stream spool, OfdPackageLoadOptions options,
        OfdModelMaterialization material, CancellationToken token,
        OfdOwnedReadAcquisition? ownedAcquisition = null, OfdOwnedZipCensus? ownedCensus = null)
    {
        material.AssertReadActive(token);
        if (spool is null) throw new ArgumentNullException(nameof(spool));
        ValidateOptions(options);
        if (Environment.Version.Major != 10) throw new NotSupportedException("Controlled ZIP accounting is audited for the .NET 10 runtime.");
        if (!spool.CanRead || !spool.CanSeek || spool.Position != 0)
            throw new NotSupportedException("Controlled OFD reads require a stable owned readable seekable spool at position zero.");
        var length = spool.Length;
        if (length < 22) throw new InvalidDataException("Controlled ZIP has no complete 22-byte EOCD footer.");
        if (length > options.MaxInputBytes) throw new InvalidDataException("OFD input exceeds the configured compressed input size limit.");
        material.ChargeCopy(8192, token); // fixed screen, options/stream wrappers and one EOF probe
        using var source = new AccountedReadStream(spool, length, material, token);
        var footer = new byte[22];
        source.Position = length - footer.Length;
        var footerRead = 0;
        while (footerRead < footer.Length)
        {
            var read = await source.ReadAsync(footer, footerRead, footer.Length - footerRead, token).ConfigureAwait(false);
            if (read == 0) throw new InvalidDataException("Truncated controlled ZIP footer.");
            footerRead += read;
        }
        if (UInt32(footer, 0) != 0x06054b50) throw new NotSupportedException("Controlled ZIP requires a final EOCD without comments or trailing data.");
        var disk = UInt16(footer, 4); var centralDisk = UInt16(footer, 6);
        var diskCount = UInt16(footer, 8); var count = UInt16(footer, 10);
        var centralSize = UInt32(footer, 12); var centralOffset = UInt32(footer, 16); var commentLength = UInt16(footer, 20);
        if (disk == ushort.MaxValue || centralDisk == ushort.MaxValue || diskCount == ushort.MaxValue || count == ushort.MaxValue ||
            centralSize == uint.MaxValue || centralOffset == uint.MaxValue)
            throw new NotSupportedException("Controlled ZIP does not support ZIP64 EOCD sentinels.");
        if (disk != 0 || centralDisk != 0 || diskCount != count || commentLength != 0)
            throw new NotSupportedException("Controlled ZIP requires a single disk, consistent counts and no archive comment.");
        if (count > options.MaxEntryCount) throw new InvalidDataException("Controlled ZIP entry count exceeds the configured limit.");
        if (checked((long)centralOffset + centralSize) != length - 22 || checked(46L * count) > centralSize)
            throw new InvalidDataException("Controlled ZIP central directory range/count is inconsistent.");
        var suffix = checked(length - centralOffset);
        // Runtime basis: 46-byte central header and 4-byte generic extra header.
        // Includes one pending malformed field (1MiB), entry/header/table shells,
        // name/comment/extra byte+UTF16 copies, and transient generic-extra objects.
        // Reserved workload allowances, never measured allocations or an RSS bound.
        var physicalRecords = checked(suffix / 46 + 1);
        var genericExtras = checked(suffix / 4 + 1);
        var metadataAllowance = checked(1048576L + 2048L * physicalRecords + 16L * suffix + 128L * genericExtras);
        material.AdmitReadMetadata(metadataAllowance, token);
        source.Position = 0;
        using var zip = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        var entries = zip.Entries;
        if (entries.Count != count || entries.Count > options.MaxEntryCount) throw new InvalidDataException("Controlled ZIP actual central entry count is inconsistent.");
        ownedCensus?.SetExpectedPhysicalCount(entries.Count, token);
        material.ReserveReadModel(checked(512L + entries.Count * 160L), token);
        var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var probe = new byte[1]; long total = 0;
        foreach (var entry in entries)
        {
            material.AssertReadActive(token);
            // entry.Name can allocate; reserve its position before accessing it.
            material.ChargeCopy(checked(128L + 4L * entry.FullName.Length), token);
            if (string.IsNullOrWhiteSpace(entry.Name))
            {
                if (ownedAcquisition is not null)
                    OfdOwnedSaveAuthority.RecordActualLoadedEntry(ownedAcquisition, entry.FullName,
                        OfdOwnedPhysicalEntryKind.Directory, entry.CompressedLength, entry.Length, null, token);
                ownedCensus?.Record(entry.FullName, null, entry.CompressedLength, entry.Length, true, null, token);
                continue;
            }
            material.ChargeCopy(checked(256L + 16L * entry.FullName.Length), token);
            var nameStorage = material.ReserveReadStorage(checked(2L * entry.FullName.Length), token);
            var name = NormalizeAndValidate(entry.FullName); material.AdoptReadString(name, nameStorage, token);
            if (result.ContainsKey(name)) throw new InvalidDataException("OFD package contains a duplicate entry: " + name);
            var size = entry.Length;
            if (size < 0 || size > int.MaxValue || size > options.MaxEntryUncompressedBytes)
                throw new InvalidDataException("OFD entry exceeds the supported array/uncompressed size limit.");
            total = checked(total + size);
            if (total > options.MaxTotalUncompressedBytes) throw new InvalidDataException("OFD package exceeds the configured total uncompressed size limit.");
            if (entry.CompressedLength > 0 && size / (double)entry.CompressedLength > options.MaxCompressionRatio)
                throw new InvalidDataException("OFD entry exceeds the configured compression ratio limit.");
            material.ChargeCopy(16384, token); // per-entry Read-mode stream/decoder working positions
            var bytes = material.AllocateReadEntry((int)size, token);
            using var input = entry.Open(); var offset = 0;
            while (offset < bytes.Length)
            {
                material.AssertReadActive(token); var requested = Math.Min(4096, bytes.Length - offset);
                material.ChargeWork(requested, token);
                var read = await input.ReadAsync(bytes, offset, requested, token).ConfigureAwait(false);
                if (read == 0) throw new InvalidDataException("OFD entry has an inconsistent expanded length.");
                offset += read;
            }
            material.ChargeWork(1, token);
            if (await input.ReadAsync(probe, 0, 1, token).ConfigureAwait(false) != 0)
                throw new InvalidDataException("OFD entry exceeds its declared expanded length.");
            material.AssertReadActive(token); material.ReserveReadModel(160, token); result.Add(name, bytes);
            if (ownedAcquisition is not null)
                OfdOwnedSaveAuthority.RecordActualLoadedEntry(ownedAcquisition, entry.FullName,
                    OfdOwnedPhysicalEntryKind.AcceptedFile, entry.CompressedLength, bytes.LongLength, bytes, token);
            ownedCensus?.Record(entry.FullName, name, entry.CompressedLength, bytes.LongLength, false, bytes, token);
        }
        if (ownedAcquisition is not null) OfdOwnedSaveAuthority.SealActualLoaderAcquisition(ownedAcquisition, token);
        ownedCensus?.Seal(token);
        material.AssertReadActive(token); return new OfdPackageArchive(result);
    }

    private static ushort UInt16(byte[] bytes, int index) => (ushort)(bytes[index] | bytes[index + 1] << 8);
    private static uint UInt32(byte[] bytes, int index) => (uint)(bytes[index] | bytes[index + 1] << 8 | bytes[index + 2] << 16 | bytes[index + 3] << 24);

    // ZipArchive's synchronous metadata reads get cancellation/work admission too.
    // The source remains caller-owned; disposal closes only this delegating shell.
    private sealed class AccountedReadStream(Stream source, long length, OfdModelMaterialization material, CancellationToken operationToken) : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length { get { Check(); return length; } }
        public override long Position { get { Check(); return source.Position; } set => Seek(value, SeekOrigin.Begin); }
        private void Check()
        {
            material.AssertReadActive(operationToken);
            if (source.Length != length || source.Position < 0 || source.Position > length) throw new InvalidDataException("Controlled ZIP spool changed length or position.");
        }
        public override int Read(byte[] buffer, int offset, int count)
        { Check(); material.ChargeWork(count, operationToken); return source.Read(buffer, offset, count); }
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
        {
            Check(); token.ThrowIfCancellationRequested(); material.ChargeWork(count, operationToken);
            var result = await source.ReadAsync(buffer, offset, count, token).ConfigureAwait(false); Check(); return result;
        }
        public override long Seek(long offset, SeekOrigin origin)
        {
            Check(); var target = origin switch { SeekOrigin.Begin => offset, SeekOrigin.Current => checked(source.Position + offset), SeekOrigin.End => checked(length + offset), _ => throw new ArgumentOutOfRangeException(nameof(origin)) };
            if (target < 0 || target > length) throw new InvalidDataException("Controlled ZIP seek leaves the admitted spool.");
            return source.Seek(target, SeekOrigin.Begin);
        }
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
