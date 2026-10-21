using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>Audited result of one independent, limited archival copy operation.</summary>
public sealed class ArchivalCopyResult
{
    internal ArchivalCopyResult(string output, string audit, string before, string after,
        IReadOnlyList<string> removed, ArchivalRuleReport report)
    { OutputPath = output; AuditPath = audit; InputSha256 = before; OutputSha256 = after; RemovedDeclarations = removed; Postflight = report; }
    /// <summary>New output package; the input file is never overwritten.</summary>
    public string OutputPath { get; }
    /// <summary>XML change manifest with per-entry payload hashes.</summary>
    public string AuditPath { get; }
    /// <summary>Compressed input package SHA-256.</summary>
    public string InputSha256 { get; }
    /// <summary>Compressed output package SHA-256.</summary>
    public string OutputSha256 { get; }
    /// <summary>Declarations deliberately removed from Document.xml.</summary>
    public IReadOnlyList<string> RemovedDeclarations { get; }
    /// <summary>Post-conversion read-only report; its Pass is only for the implemented subset.</summary>
    public ArchivalRuleReport Postflight { get; }
    /// <summary>Signed inputs are rejected; a successful run did not invalidate a signature.</summary>
    public bool SignaturesInvalidated => false;
}

/// <summary>Opt-in 6.2.2(a)-(b) copy handler. It refuses opaque targets and signed input.</summary>
public sealed class ArchivalCopyConverter
{
    /// <summary>Creates a distinct OFD and XML audit sidecar after complete postflight verification.</summary>
    public async Task<ArchivalCopyResult> ConvertAsync(string inputPath, string outputPath,
        ArchivalPreflightOptions options, CancellationToken cancellationToken = default)
    {
        if (inputPath == null) throw new ArgumentNullException(nameof(inputPath));
        if (outputPath == null) throw new ArgumentNullException(nameof(outputPath));
        if (options == null) throw new ArgumentNullException(nameof(options));
        var source = Path.GetFullPath(inputPath);
        var output = Path.GetFullPath(outputPath);
        var audit = output + ".audit.xml";
        if (string.Equals(source, output, StringComparison.OrdinalIgnoreCase) || File.Exists(output) || File.Exists(audit))
            throw new IOException("Output and audit paths must be new and distinct from the input.");
        var snapshotOptions = options.Snapshot();
        byte[] sourceBytes;
        using (var file = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            if (file.Length > snapshotOptions.LoadOptions.MaxInputBytes || file.Length > int.MaxValue)
                throw new InvalidDataException("Input exceeds the configured compressed input budget.");
            sourceBytes = new byte[checked((int)file.Length)];
            var offset = 0;
            while (offset < sourceBytes.Length)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var read = await file.ReadAsync(sourceBytes, offset, sourceBytes.Length - offset, cancellationToken).ConfigureAwait(false);
                if (read == 0) throw new IOException("Input changed during snapshot acquisition.");
                offset += read;
            }
            if (file.ReadByte() != -1) throw new IOException("Input grew during snapshot acquisition.");
        }
        var beforeSha = Hash(sourceBytes);
        ArchivalRuleReport before;
        OfdPackageArchive original;
        using (var stream = new MemoryStream(sourceBytes, false))
        {
            before = await new ArchivalRuleService().CheckAsync(stream, snapshotOptions, cancellationToken).ConfigureAwait(false);
            stream.Position = 0;
            original = await new OfdPackageLoader().LoadAsync(stream, snapshotOptions.LoadOptions, cancellationToken).ConfigureAwait(false);
        }
        if (before.LocalPreflight.GetRuleOutcome("APF-006") != PreflightOutcome.NotApplicable ||
            original.EntryNames.Any(x => x.IndexOf("/Signs/", StringComparison.OrdinalIgnoreCase) >= 0 ||
                x.EndsWith("Signatures.xml", StringComparison.OrdinalIgnoreCase) ||
                x.EndsWith("SignedValue.dat", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException("SignedInputRequiresResigning: this handler refuses to invalidate or silently clear signatures.");
        if (before.LocalPreflight.RunState != PreflightRunState.Complete ||
            before.LocalPreflight.RuleResults.Any(x => x.Outcome == PreflightOutcome.Fail || x.Outcome == PreflightOutcome.Unknown) ||
            before.RuleResults.Any(x => x.Outcome == PreflightOutcome.Unknown) || before.DocumentEntry == null)
            throw new InvalidDataException("Preflight is incomplete, failed, or uncertain; no copy was produced.");

        var document = ArchivalRuleService.Read(original, before.DocumentEntry, snapshotOptions)
            ?? throw new InvalidDataException("Document XML cannot be read safely.");
        var root = document.Root ?? throw new InvalidDataException("Document XML has no root.");
        var removed = new List<string>();
        foreach (var handler in ArchivalRuleHandlers.Supported)
        {
            if (handler.Apply(root)) removed.Add(handler.Rule.Element);
        }
        if (removed.Count == 0) throw new InvalidOperationException("Input already satisfies the selected rules; no copy is needed.");
        var changed = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in original.EntryNames) changed.Add(entry, original.GetBytes(entry));
        changed[before.DocumentEntry] = Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));

        var directory = Path.GetDirectoryName(output) ?? throw new IOException("Output has no directory.");
        if (!Directory.Exists(directory)) throw new DirectoryNotFoundException(directory);
        var temporary = Path.Combine(directory, "." + Path.GetFileName(output) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        var temporaryAudit = temporary + ".audit.xml";
        var published = false;
        try
        {
            // A stored input entry can be highly compressible. Recompressing it may
            // violate the caller's ratio budget even though the source was admissible.
            // Relax only offending entries, and verify the actual ZIP metadata each time.
            var compression = new Dictionary<string, CompressionLevel>(StringComparer.OrdinalIgnoreCase);
            for (var attempt = 0; attempt < 3; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (attempt != 0) File.Delete(temporary);
                WritePackage(temporary, changed, compression, cancellationToken);
                var offenders = EntriesOverRatio(temporary, snapshotOptions.LoadOptions.MaxCompressionRatio);
                if (offenders.Count == 0) break;
                if (attempt == 2) throw new InvalidDataException("Repacked entries exceed the configured compression ratio limit.");
                foreach (var name in offenders)
                    compression[name] = attempt == 0 ? CompressionLevel.Fastest : CompressionLevel.NoCompression;
            }
            byte[] outputBytes = File.ReadAllBytes(temporary);
            ArchivalRuleReport after;
            OfdPackageArchive verified;
            using (var stream = new MemoryStream(outputBytes, false))
            {
                after = await new ArchivalRuleService().CheckAsync(stream, snapshotOptions, cancellationToken).ConfigureAwait(false);
                stream.Position = 0;
                verified = await new OfdPackageLoader().LoadAsync(stream, snapshotOptions.LoadOptions, cancellationToken).ConfigureAwait(false);
            }
            if (after.SummaryOutcome != PreflightOutcome.Pass || after.LocalPreflight.RunState != PreflightRunState.Complete ||
                after.LocalPreflight.RuleResults.Any(x => x.Outcome == PreflightOutcome.Fail || x.Outcome == PreflightOutcome.Unknown) ||
                !original.EntryNames.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).SequenceEqual(verified.EntryNames.OrderBy(x => x, StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase) ||
                original.EntryNames.Where(x => !string.Equals(x, before.DocumentEntry, StringComparison.OrdinalIgnoreCase))
                    .Any(x => !original.GetBytes(x).SequenceEqual(verified.GetBytes(x))) ||
                !changed[before.DocumentEntry].SequenceEqual(verified.GetBytes(before.DocumentEntry)))
                throw new InvalidDataException("Postflight or unchanged-entry verification failed.");
            using (var file = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
                if (Hash(file) != beforeSha) throw new IOException("Input changed after snapshot acquisition.");
            var afterSha = Hash(outputBytes);
            var entryRecords = original.EntryNames.OrderBy(x => x, StringComparer.Ordinal).Select(name =>
                new XElement("Entry", new XAttribute("path", name), new XAttribute("beforeSha256", Hash(original.GetBytes(name))),
                    new XAttribute("afterSha256", Hash(verified.GetBytes(name))),
                    new XAttribute("changed", string.Equals(name, before.DocumentEntry, StringComparison.OrdinalIgnoreCase))));
            var removedRecords = removed.Select(name => new XElement("Declaration", new XAttribute("name", name),
                new XAttribute("clause", ArchivalRuleCatalog.Rules.Single(r => r.Element == name).Clause),
                new XAttribute("informationLoss", "Permission or view preference declaration removed")));
            var manifest = new XDocument(new XElement("ArchivalCopyAudit",
                new XAttribute("standard", "GB/T 42133-2022"), new XAttribute("scope", "6.2.2(a)-(b) only"),
                new XAttribute("sourceSha256", "7d76c086e1384ff200c0df8f5a40bf67660e841340c5e2630e582c81c22f1b82"),
                new XAttribute("certification", "NotAssessed"), new XAttribute("signatureDisposition", "UnsignedInputOnly"),
                new XElement("Input", new XAttribute("sha256", beforeSha), new XAttribute("path", source)),
                new XElement("Output", new XAttribute("sha256", afterSha), new XAttribute("path", output)),
                new XElement("Reinspection", new XAttribute("inputLocal", before.LocalPreflight.SummaryOutcome),
                    new XAttribute("outputLocal", after.LocalPreflight.SummaryOutcome),
                    new XAttribute("inputSubset", before.SummaryOutcome), new XAttribute("outputSubset", after.SummaryOutcome),
                    after.RuleResults.Select(result => new XElement("Rule", new XAttribute("id", result.Rule.RuleId),
                        new XAttribute("clause", result.Rule.Clause), new XAttribute("outcome", result.Outcome),
                        new XAttribute("evidenceEntry", result.Evidence.Entry),
                        new XAttribute("evidenceSelector", result.Evidence.Selector)))),
                new XElement("RemovedDeclarations", removedRecords),
                new XElement("Entries", entryRecords)));
            using (var file = new FileStream(temporaryAudit, FileMode.CreateNew, FileAccess.Write, FileShare.None)) manifest.Save(file);
            cancellationToken.ThrowIfCancellationRequested();
            if (File.Exists(output) || File.Exists(audit)) throw new IOException("Output path was claimed before publication.");
            File.Move(temporary, output);
            try { File.Move(temporaryAudit, audit); }
            catch { File.Delete(output); throw; }
            published = true;
            return new ArchivalCopyResult(output, audit, beforeSha, afterSha, removed.AsReadOnly(), after);
        }
        finally
        {
            if (!published) { if (File.Exists(temporary)) File.Delete(temporary); if (File.Exists(temporaryAudit)) File.Delete(temporaryAudit); }
        }
    }

    private static string Hash(byte[] bytes)
    { using var sha = SHA256.Create(); return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }
    private static string Hash(Stream stream)
    { using var sha = SHA256.Create(); return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant(); }

    private static void WritePackage(string path, IReadOnlyDictionary<string, byte[]> entries,
        IReadOnlyDictionary<string, CompressionLevel> compression, CancellationToken cancellationToken)
    {
        using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        using var zip = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false);
        foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var level = compression.TryGetValue(entry.Key, out var chosen) ? chosen : CompressionLevel.Optimal;
            var item = zip.CreateEntry(entry.Key, level);
            item.LastWriteTime = new DateTimeOffset(2020, 1, 1, 0, 0, 0, TimeSpan.Zero);
            using var target = item.Open();
            target.Write(entry.Value, 0, entry.Value.Length);
        }
    }

    private static IReadOnlyList<string> EntriesOverRatio(string path, double limit)
    {
        using var file = File.OpenRead(path);
        using var zip = new ZipArchive(file, ZipArchiveMode.Read);
        return zip.Entries.Where(entry => entry.CompressedLength > 0 &&
            entry.Length / (double)entry.CompressedLength > limit).Select(entry => entry.FullName).ToArray();
    }
}
