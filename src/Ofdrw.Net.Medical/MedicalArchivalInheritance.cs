using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Medical;

/// <summary>How a GB/T 42133-2022 ledger item relates to an OFD-H document under GB/T 48666-2026 clause 6.1.2(b).</summary>
public enum MedicalArchivalDisposition
{
    /// <summary>
    /// Treated as a "relevant requirement" of 6.1.2(b). The reading is <see cref="MedicalMatrixConfidence.Unverified"/>:
    /// the standard does not list the relevant clauses. A mapped Fail makes the medical standard outcome Fail.
    /// </summary>
    Inherited,
    /// <summary>GB/T 48666 states its own, different rule for the same subject; the archival outcome is informational.</summary>
    Superseded,
    /// <summary>Applies when the medical record is assembled into an archival OFD (7.3(g), 8(b), 9.1(c)), not to OFD-H itself.</summary>
    ArchivalStageOnly
}

/// <summary>One row of the GB/T 42133 inheritance ledger.</summary>
public sealed class MedicalArchivalInheritanceEntry
{
    internal MedicalArchivalInheritanceEntry(string clause, MedicalArchivalDisposition disposition, string basis, string reason)
    { ArchivalClause = clause; Disposition = disposition; MedicalBasis = basis; Reason = reason; }
    /// <summary>GB/T 42133-2022 ledger clause, exactly as in <see cref="ArchivalClauseCatalog"/>.</summary>
    public string ArchivalClause { get; }
    /// <summary>Disposition under GB/T 48666.</summary>
    public MedicalArchivalDisposition Disposition { get; }
    /// <summary>The GB/T 48666 clause that decides the disposition.</summary>
    public string MedicalBasis { get; }
    /// <summary>Short explanation of the disposition.</summary>
    public string Reason { get; }
    /// <summary>Unverified for inherited rows; the explicit rows quote a GB/T 48666 clause.</summary>
    public MedicalMatrixConfidence Confidence =>
        Disposition == MedicalArchivalDisposition.Inherited ? MedicalMatrixConfidence.Unverified : MedicalMatrixConfidence.Verified;
}

/// <summary>
/// Ledger that maps every GB/T 42133-2022 item to an OFD-H disposition. Only the explicit rows below differ from
/// <see cref="MedicalArchivalDisposition.Inherited"/>; the baseline document explains each row.
/// </summary>
public static class MedicalArchivalInheritance
{
    /// <summary>Ledger revision; bump when any disposition changes.</summary>
    public const string Revision = "gbt48666-2026-archival-inheritance/1";

    private const string Actions = "7.3(f) requires internal jump actions and external links for associated objects; GB/T 42133 removes non-Goto and external actions.";
    private const string ArchiveStage = "Archive-time processing: 7.3(g) assembles archival OFD files and 9.1(c) applies GB/T 42133 processing then.";
    private const string Attachments = "7.4 and 7.2(e) keep internal and external attachments, including audio, video and DICOM.";

    private static readonly Dictionary<string, (MedicalArchivalDisposition, string, string)> Explicit =
        new Dictionary<string, (MedicalArchivalDisposition, string, string)>(StringComparer.Ordinal)
        {
            ["6.2.1(a)"] = (MedicalArchivalDisposition.Superseded, "7.1(a)", "OFD-H declares DocType=OFD-H, not OFD-A."),
            ["6.2.2(a)"] = (MedicalArchivalDisposition.Superseded, "5.3.2(c)", "Exchanged documents need permission control; GB/T 42133 removes Permissions."),
            ["6.2.2(c)"] = (MedicalArchivalDisposition.Superseded, "7.3(f)", Actions),
            ["6.2.3(c)"] = (MedicalArchivalDisposition.Superseded, "7.3(f)", Actions),
            ["6.2.5(a)"] = (MedicalArchivalDisposition.Superseded, "7.3(f)", Actions),
            ["6.3.3(b)"] = (MedicalArchivalDisposition.Superseded, "7.3(f)", Actions),
            ["6.9(a)"] = (MedicalArchivalDisposition.Superseded, "7.3(f)", Actions),
            ["6.9(b)"] = (MedicalArchivalDisposition.Superseded, "7.3(f)", Actions),
            ["6.2.6(e)"] = (MedicalArchivalDisposition.Superseded, "7.2(d)", "7.2(d) lists its own permitted image formats."),
            ["6.2.6(g)"] = (MedicalArchivalDisposition.Superseded, "7.2(f)", "7.2(f) and 7.2(g) allow embedded MP3/MP4 audio and video."),
            ["6.7"] = (MedicalArchivalDisposition.Superseded, "7.2(g)", "7.2(g) allows embedded video."),
            ["6.12(a)"] = (MedicalArchivalDisposition.Superseded, "6.2(a)", "Table 1 keeps Extensions.xml when extensions are described."),
            ["6.12(b)"] = (MedicalArchivalDisposition.Superseded, "6.2(a)", "Table 1 keeps Extensions.xml when extensions are described."),
            ["6.14"] = (MedicalArchivalDisposition.Superseded, "7.6", "7.6 keeps historical versions or stores their changes as internal attachments."),
            ["6.16"] = (MedicalArchivalDisposition.Superseded, "8(b)", "Temporary files may be encrypted; decryption happens before archiving or migration."),
            ["6.2.1(d)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "7.3(g)", ArchiveStage),
            ["6.2.1(e)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "7.3(g)", ArchiveStage),
            ["6.2.2(b)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "7.3(g)", ArchiveStage),
            ["6.2.2(d)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "7.4", Attachments),
            ["6.2.2(e)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "7.3(g)", ArchiveStage),
            ["6.10.1(a)(1)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "6.3(f)", "Signed OFD-H files keep accepting annotations; freezing them is an archive step."),
            ["6.10.1(a)(2)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "6.3(f)", "Signed OFD-H files keep accepting annotations; freezing them is an archive step."),
            ["6.10.1(a)(3)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "6.3(f)", "Signed OFD-H files keep accepting annotations; freezing them is an archive step."),
            ["6.10.1(b)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "6.3(f)", "Signed OFD-H files keep accepting annotations; freezing them is an archive step."),
            ["6.10.2"] = (MedicalArchivalDisposition.ArchivalStageOnly, "6.3(f)", "Signed OFD-H files keep accepting annotations; freezing them is an archive step."),
            ["6.13.1(c)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "7.3(g)", ArchiveStage),
            ["6.13.2(a)(1)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "8(c)", "OFD-H keeps live signatures; de-technicalizing them is an archive step."),
            ["6.13.2(a)(2)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "8(c)", "OFD-H keeps live signatures; de-technicalizing them is an archive step."),
            ["6.13.2(b)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "8(c)", "OFD-H keeps live signatures; de-technicalizing them is an archive step."),
            ["6.13.2(c)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "8(c)", "OFD-H keeps live signatures; de-technicalizing them is an archive step."),
            ["6.13.3(a)(1)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "8(c)", "OFD-H keeps live seals; de-technicalizing them is an archive step."),
            ["6.13.3(a)(2)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "8(c)", "OFD-H keeps live seals; de-technicalizing them is an archive step."),
            ["6.13.3(b)"] = (MedicalArchivalDisposition.ArchivalStageOnly, "8(c)", "OFD-H keeps live seals; de-technicalizing them is an archive step."),
            ["7.1(1)"] = (MedicalArchivalDisposition.Superseded, "9.1", "Chapter 9 states the medical generator requirements."),
            ["7.1(2)"] = (MedicalArchivalDisposition.Superseded, "9.1", "Chapter 9 states the medical generator requirements."),
            ["7.2(1)"] = (MedicalArchivalDisposition.Superseded, "9.2", "Chapter 9 states the medical reader requirements."),
            ["7.2(2)"] = (MedicalArchivalDisposition.Superseded, "9.2(b)", "9.2(b) locks page editing for OFD-H instead of OFD-A."),
        };

    /// <summary>One row per <see cref="ArchivalClauseCatalog"/> item, in ledger order.</summary>
    public static IReadOnlyList<MedicalArchivalInheritanceEntry> Entries { get; } = Array.AsReadOnly(
        ArchivalClauseCatalog.Clauses.Select(definition => Create(definition.Clause)).ToArray());

    /// <summary>Returns the row for an exact GB/T 42133 ledger clause.</summary>
    /// <param name="archivalClause">Ledger clause identifier.</param>
    /// <returns>The inheritance row.</returns>
    /// <exception cref="ArgumentException">The clause is not in the GB/T 42133 ledger.</exception>
    public static MedicalArchivalInheritanceEntry Get(string archivalClause) =>
        Entries.FirstOrDefault(x => x.ArchivalClause == archivalClause) ?? throw new ArgumentException("Unknown GB/T 42133 ledger clause.", nameof(archivalClause));

    private static MedicalArchivalInheritanceEntry Create(string clause)
    {
        if (Explicit.TryGetValue(clause, out var row)) return new MedicalArchivalInheritanceEntry(clause, row.Item1, row.Item2, row.Item3);
        if (clause.StartsWith("A.1/", StringComparison.Ordinal))
            return new MedicalArchivalInheritanceEntry(clause, MedicalArchivalDisposition.Superseded, "7.4", Attachments);
        if (clause.StartsWith("6.15", StringComparison.Ordinal))
            return new MedicalArchivalInheritanceEntry(clause, MedicalArchivalDisposition.Superseded, "7.4", Attachments);
        if (clause.StartsWith("5.", StringComparison.Ordinal))
            return new MedicalArchivalInheritanceEntry(clause, MedicalArchivalDisposition.ArchivalStageOnly, "5.1", "GB/T 42133 scope and archival-unit organization; GB/T 48666 chapter 5 states its own scenarios.");
        return new MedicalArchivalInheritanceEntry(clause, MedicalArchivalDisposition.Inherited, "6.1.2(b)", "Relevant GB/T 42133 requirement (reading unverified).");
    }

    internal static IEnumerable<string> ExplicitClauses => Explicit.Keys;
}
