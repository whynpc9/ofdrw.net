using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging;

/// <summary>
/// Record of one ordinary-OFD derivation. The derivative is a new, ordinary OFD: it does not inherit
/// any archival conclusion, report or signature validity of its source.
/// </summary>
public sealed class OfdOrdinaryDerivationResult
{
    internal OfdOrdinaryDerivationResult(string sourcePath, string sourceSha256, long sourceLength, string sourceDocType,
        string outputPath, string outputSha256, long outputLength, string changedEntry)
    {
        SourcePath = sourcePath; SourceSha256 = sourceSha256; SourceLength = sourceLength; SourceDocType = sourceDocType;
        OutputPath = outputPath; OutputSha256 = outputSha256; OutputLength = outputLength; ChangedEntry = changedEntry;
    }

    /// <summary>The operation name recorded for the derivation, <see cref="OfdOrdinaryDerivation.OperationName"/>.</summary>
    public string Operation => OfdOrdinaryDerivation.OperationName;
    /// <summary>Full path of the archival original; it is never modified.</summary>
    public string SourcePath { get; }
    /// <summary>Lowercase hexadecimal SHA-256 of the source file bytes that were snapshotted.</summary>
    public string SourceSha256 { get; }
    /// <summary>Source file length in bytes.</summary>
    public long SourceLength { get; }
    /// <summary>The root DocType declared by the source, as read from OFD.xml (always <c>OFD-A</c> on success).</summary>
    public string SourceDocType { get; }
    /// <summary>Full path of the newly created derivative.</summary>
    public string OutputPath { get; }
    /// <summary>Lowercase hexadecimal SHA-256 of the derivative file.</summary>
    public string OutputSha256 { get; }
    /// <summary>Derivative file length in bytes.</summary>
    public long OutputLength { get; }
    /// <summary>The root DocType declared by the derivative, always <c>OFD</c>.</summary>
    public string OutputDocType => OfdDocTypes.Ordinary;
    /// <summary>The only package entry whose content differs, <c>OFD.xml</c>; only its root DocType attribute value changed.</summary>
    public string ChangedEntry { get; }
    /// <summary>Always <see langword="false"/>: the derivative carries no archival conclusion of its source.</summary>
    public bool InheritsArchivalConclusion => false;
}

/// <summary>
/// Creates an ordinary-OFD derivative of an archival original (an input whose OFD.xml declares
/// <c>DocType="OFD-A"</c>) so that it can be edited. The derivative is a different, new file whose root
/// DocType is rewritten to <c>OFD</c>; every other package entry keeps byte-identical content and OFD.xml
/// differs only in that attribute value. Signed and encrypted inputs are refused rather than stripped.
/// </summary>
public static class OfdOrdinaryDerivation
{
    /// <summary>The operation name recorded in <see cref="OfdOrdinaryDerivationResult.Operation"/>.</summary>
    public const string OperationName = "DeriveOrdinaryOfd";

    private const string RootEntry = "OFD.xml";
    private static readonly Regex DocTypeAttribute = new(@"(?<=\s)DocType\s*=\s*", RegexOptions.CultureInvariant);

    /// <summary>
    /// Derives an ordinary OFD from an archival original into <paramref name="outputPath"/>, which must not exist.
    /// </summary>
    /// <param name="inputPath">The archival original; read once into a bounded snapshot and never modified.</param>
    /// <param name="outputPath">A new file path distinct from the input; it must not exist.</param>
    /// <param name="loadOptions">Package budgets; defaults to <see cref="OfdPackageLoadOptions"/> defaults.</param>
    /// <param name="cancellationToken">Cancellation; a canceled run publishes nothing.</param>
    /// <returns>The derivation record with source and output identities.</returns>
    /// <exception cref="InvalidOperationException">The input does not declare OFD-A and needs no derivation.</exception>
    /// <exception cref="NotSupportedException">The input is signed or encrypted, or its declaration cannot be rewritten exactly.</exception>
    /// <exception cref="IOException">The output exists, equals the input, or the input changed during the run.</exception>
    /// <remarks>On any failure or cancellation no output file is left behind and the source is untouched.</remarks>
    public static async Task<OfdOrdinaryDerivationResult> CreateAsync(string inputPath, string outputPath,
        OfdPackageLoadOptions? loadOptions = null, CancellationToken cancellationToken = default)
    {
        if (inputPath is null) throw new ArgumentNullException(nameof(inputPath));
        if (outputPath is null) throw new ArgumentNullException(nameof(outputPath));
        var options = loadOptions ?? new OfdPackageLoadOptions();
        cancellationToken.ThrowIfCancellationRequested();
        var source = Path.GetFullPath(inputPath);
        var output = Path.GetFullPath(outputPath);
        if (string.Equals(source, output, StringComparison.Ordinal))
            throw new IOException("The derivative must be written to a different file than the archival original.");
        // File.Exists follows the destination volume's own aliasing (case folding, links), so a
        // pre-existing output, including any alias of the input, is refused before reading.
        if (File.Exists(output) || Directory.Exists(output))
            throw new IOException("The derivative output path already exists; choose a new file. Existing files are never overwritten.");
        var directory = Path.GetDirectoryName(output) ?? throw new IOException("The output path has no directory.");
        if (!Directory.Exists(directory)) throw new DirectoryNotFoundException(directory);

        var sourceBytes = await SnapshotAsync(source, options.MaxInputBytes, cancellationToken).ConfigureAwait(false);
        var sourceSha = Hash(sourceBytes);
        OfdPackageArchive original;
        using (var stream = new MemoryStream(sourceBytes, false))
            original = await new OfdPackageLoader().LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
        var rootName = original.EntryNames.SingleOrDefault(name => string.Equals(name, RootEntry, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidDataException("OFD.xml is missing.");
        var rootBytes = original.GetBytes(rootName);
        var rootDocument = ParseXml(rootBytes);
        var root = rootDocument.Root;
        if (root is null || root.Name.LocalName != "OFD") throw new InvalidDataException("OFD.xml has no OFD root element.");
        var declared = root.Attribute("DocType")?.Value;
        if (!OfdDocTypes.IsArchival(declared))
            throw new InvalidOperationException(
                $"The input does not declare DocType=\"OFD-A\" (declared: {(declared is null ? "none" : "\"" + declared + "\"")}); " +
                "it is not an archival original and needs no derivation. No output was written.");
        RejectSignedOrEncrypted(original, root);

        var derivedRoot = RewriteDocType(rootBytes);
        var expected = new XDocument(rootDocument);
        expected.Root!.SetAttributeValue("DocType", OfdDocTypes.Ordinary);
        if (!XNode.DeepEquals(expected, ParseXml(derivedRoot)))
            throw new NotSupportedException("The DocType declaration could not be rewritten without changing other OFD.xml content.");

        var order = EntryOrder(sourceBytes, original);
        var temporary = Path.Combine(directory, "." + Path.GetFileName(output) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        var published = false;
        try
        {
            var stored = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var attempt = 0; ; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (attempt != 0) File.Delete(temporary);
                WritePackage(temporary, order, original, rootName, derivedRoot, stored, cancellationToken);
                var offenders = EntriesOverRatio(temporary, options.MaxCompressionRatio);
                if (offenders.Count == 0) break;
                if (attempt == 1) throw new InvalidDataException("Repacked entries exceed the configured compression ratio limit.");
                foreach (var name in offenders) stored.Add(name);
            }

            var outputBytes = File.ReadAllBytes(temporary);
            OfdPackageArchive verified;
            using (var stream = new MemoryStream(outputBytes, false))
                verified = await new OfdPackageLoader().LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
            var originalNames = original.EntryNames.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();
            var verifiedNames = verified.EntryNames.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();
            if (!originalNames.SequenceEqual(verifiedNames, StringComparer.Ordinal) ||
                originalNames.Where(name => !string.Equals(name, rootName, StringComparison.Ordinal))
                    .Any(name => !original.GetBytes(name).SequenceEqual(verified.GetBytes(name))) ||
                !verified.GetBytes(rootName).SequenceEqual(derivedRoot) ||
                ParseXml(verified.GetBytes(rootName)).Root?.Attribute("DocType")?.Value != OfdDocTypes.Ordinary)
                throw new InvalidDataException("Post-write verification of the derivative failed; no output was published.");
            using (var file = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
                if (Hash(file) != sourceSha) throw new IOException("The archival original changed during derivation; no output was published.");
            cancellationToken.ThrowIfCancellationRequested();
            if (File.Exists(output)) throw new IOException("The output path was claimed before publication.");
            File.Move(temporary, output);
            published = true;
            return new OfdOrdinaryDerivationResult(source, sourceSha, sourceBytes.LongLength, declared!,
                output, Hash(outputBytes), outputBytes.LongLength, rootName);
        }
        finally
        {
            if (!published && File.Exists(temporary)) File.Delete(temporary);
        }
    }

    private static async Task<byte[]> SnapshotAsync(string path, long maxBytes, CancellationToken token)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (file.Length > maxBytes || file.Length > int.MaxValue)
            throw new InvalidDataException("Input exceeds the configured compressed input budget.");
        var bytes = new byte[checked((int)file.Length)];
        var offset = 0;
        while (offset < bytes.Length)
        {
            token.ThrowIfCancellationRequested();
            var read = await file.ReadAsync(bytes, offset, bytes.Length - offset, token).ConfigureAwait(false);
            if (read == 0) throw new IOException("Input changed during snapshot acquisition.");
            offset += read;
        }
        if (file.ReadByte() != -1) throw new IOException("Input grew during snapshot acquisition.");
        return bytes;
    }

    private static void RejectSignedOrEncrypted(OfdPackageArchive archive, XElement root)
    {
        if (root.Descendants().Any(node => node.Name.LocalName == "Signatures") ||
            archive.EntryNames.Any(name => name.Split('/').Any(part => part.Equals("Signs", StringComparison.OrdinalIgnoreCase)) ||
                name.EndsWith("Signatures.xml", StringComparison.OrdinalIgnoreCase) ||
                name.EndsWith("Signature.xml", StringComparison.OrdinalIgnoreCase) ||
                name.EndsWith("SignedValue.dat", StringComparison.OrdinalIgnoreCase)))
            throw new NotSupportedException(
                "The archival original is signed. Derivation refuses signed inputs: it does not strip signatures and cannot claim they remain valid. No output was written.");
        if (archive.EntryNames.Any(name => name.Equals("Encryptions.xml", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("OFDEntries.xml", StringComparison.OrdinalIgnoreCase) ||
                name.StartsWith("PasswordCrypto/", StringComparison.OrdinalIgnoreCase)) ||
            root.DescendantsAndSelf().Any(node => node.Name.LocalName.IndexOf("Encrypt", StringComparison.OrdinalIgnoreCase) >= 0))
            throw new NotSupportedException("The input is encrypted. Derivation refuses encrypted inputs. No output was written.");
    }

    private static XDocument ParseXml(byte[] bytes)
    {
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
        using var stream = new MemoryStream(bytes, false);
        using var reader = XmlReader.Create(stream, settings);
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    /// <summary>Replaces the literal value of the root DocType attribute in place, keeping every other byte.</summary>
    private static byte[] RewriteDocType(byte[] bytes)
    {
        var bom = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
        string text;
        try { text = new UTF8Encoding(false, true).GetString(bytes, bom, bytes.Length - bom); }
        catch (DecoderFallbackException) { throw new NotSupportedException("OFD.xml must be UTF-8 to rewrite its declaration exactly."); }
        var start = RootStart(text);
        var end = TagEnd(text, start);
        var tag = text.Substring(start, end - start);
        var matches = DocTypeAttribute.Matches(tag);
        if (matches.Count != 1) throw new NotSupportedException("OFD.xml root must carry exactly one literal DocType attribute.");
        var valueStart = start + matches[0].Index + matches[0].Length;
        var quote = text[valueStart];
        if ((quote != '"' && quote != '\'') ||
            string.CompareOrdinal(text, valueStart + 1, OfdDocTypes.Archival, 0, OfdDocTypes.Archival.Length) != 0 ||
            valueStart + 1 + OfdDocTypes.Archival.Length >= text.Length || text[valueStart + 1 + OfdDocTypes.Archival.Length] != quote)
            throw new NotSupportedException("The OFD.xml DocType value is not the literal \"OFD-A\" and cannot be rewritten exactly.");
        var byteOffset = bom + Encoding.UTF8.GetByteCount(text.Substring(0, valueStart + 1));
        var replacement = Encoding.UTF8.GetBytes(OfdDocTypes.Ordinary);
        var archival = Encoding.UTF8.GetByteCount(OfdDocTypes.Archival);
        var result = new byte[bytes.Length - archival + replacement.Length];
        Buffer.BlockCopy(bytes, 0, result, 0, byteOffset);
        Buffer.BlockCopy(replacement, 0, result, byteOffset, replacement.Length);
        Buffer.BlockCopy(bytes, byteOffset + archival, result, byteOffset + replacement.Length, bytes.Length - byteOffset - archival);
        return result;
    }

    private static int RootStart(string text)
    {
        var index = 0;
        while (true)
        {
            while (index < text.Length && char.IsWhiteSpace(text[index])) index++;
            if (index >= text.Length || text[index] != '<') throw new NotSupportedException("OFD.xml has no recognizable root element.");
            if (string.CompareOrdinal(text, index, "<?", 0, 2) == 0) index = Skip(text, index, "?>");
            else if (string.CompareOrdinal(text, index, "<!--", 0, 4) == 0) index = Skip(text, index, "-->");
            else if (string.CompareOrdinal(text, index, "<!", 0, 2) == 0) throw new NotSupportedException("OFD.xml document type declarations are not supported.");
            else return index;
        }
    }

    private static int Skip(string text, int index, string terminator)
    {
        var end = text.IndexOf(terminator, index, StringComparison.Ordinal);
        if (end < 0) throw new NotSupportedException("OFD.xml prolog is not well formed.");
        return end + terminator.Length;
    }

    private static int TagEnd(string text, int start)
    {
        char quote = '\0';
        for (var index = start + 1; index < text.Length; index++)
        {
            var current = text[index];
            if (quote != '\0') { if (current == quote) quote = '\0'; }
            else if (current is '"' or '\'') quote = current;
            else if (current == '>') return index + 1;
        }
        throw new NotSupportedException("OFD.xml root start tag is not well formed.");
    }

    private static IReadOnlyList<(string Name, DateTimeOffset Modified)> EntryOrder(byte[] sourceBytes, OfdPackageArchive archive)
    {
        var result = new List<(string, DateTimeOffset)>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (var stream = new MemoryStream(sourceBytes, false))
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
            foreach (var entry in zip.Entries)
                if (archive.Contains(entry.FullName) && seen.Add(entry.FullName)) result.Add((entry.FullName, entry.LastWriteTime));
        foreach (var name in archive.EntryNames.OrderBy(x => x, StringComparer.Ordinal))
            if (seen.Add(name)) result.Add((name, new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero)));
        return result;
    }

    private static void WritePackage(string path, IReadOnlyList<(string Name, DateTimeOffset Modified)> order, OfdPackageArchive archive,
        string rootName, byte[] derivedRoot, ISet<string> stored, CancellationToken token)
    {
        using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        using var zip = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false);
        foreach (var (name, modified) in order)
        {
            token.ThrowIfCancellationRequested();
            var data = string.Equals(name, rootName, StringComparison.Ordinal) ? derivedRoot : archive.GetBytes(name);
            var entry = zip.CreateEntry(name, stored.Contains(name) ? CompressionLevel.NoCompression : CompressionLevel.Optimal);
            entry.LastWriteTime = modified;
            using var target = entry.Open();
            target.Write(data, 0, data.Length);
        }
    }

    private static IReadOnlyList<string> EntriesOverRatio(string path, double limit)
    {
        using var file = File.OpenRead(path);
        using var zip = new ZipArchive(file, ZipArchiveMode.Read);
        return zip.Entries.Where(entry => entry.CompressedLength > 0 && entry.Length / (double)entry.CompressedLength > limit)
            .Select(entry => entry.FullName).ToArray();
    }

    private static string Hash(byte[] bytes)
    { using var sha = SHA256.Create(); return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }

    private static string Hash(Stream stream)
    { using var sha = SHA256.Create(); return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant(); }
}
