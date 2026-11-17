using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival;

/// <summary>A versioned description of the inputs an assessment claims to understand. Not a standards profile.</summary>
public sealed class ArchivalInputProfile
{
    /// <summary>Identity of the first profile.</summary>
    public const string SelfGeneratedSimpleId = "ofdrw-archival-profile/self-generated-simple/1";

    private ArchivalInputProfile(string id, string description, string[] accepted, string[] excluded)
    {
        Id = id; Description = description;
        AcceptedContent = Array.AsReadOnly(accepted); ExcludedContent = Array.AsReadOnly(excluded);
    }

    /// <summary>
    /// Simple Layout output or Native DOCX→OFD with controlled fonts: one DocBody, unsigned, unencrypted,
    /// no attachments, no Versions, no actions, annotations, extensions or composite graphics. Producer provenance
    /// is not trusted; every condition is checked against package content.
    /// </summary>
    public static ArchivalInputProfile SelfGeneratedSimple { get; } = new ArchivalInputProfile(SelfGeneratedSimpleId,
        "Single-DocBody, unsigned, unencrypted OFD without attachments, Versions, actions, annotations, extensions or composite graphics, as produced by simple Layout output or Native DOCX to OFD. Source CustomTags are retained but their reference closure is not yet verified.",
        new[]
        {
            "OFD/DocBody (exactly one) with DocInfo and DocRoot",
            "Document/CommonData, Pages, Outlines without actions, Bookmarks, CustomTags, Permissions, VPreferences",
            "Page/Area, Template, PageRes, Content",
            "Layer objects: TextObject, PathObject, ImageObject, PageBlock"
        },
        new[]
        {
            "Multiple DocBody", "DocBody/Versions", "DocBody/Signatures or signature payloads",
            "Encryptions.xml or PasswordCrypto/ entries", "Document/Attachments or attachment indexes",
            "Document/Annotations", "Document/Extensions", "Document, outline, page or object Actions",
            "CompositeObject", "Unrecognized or foreign-namespace control content"
        });

    /// <summary>Versioned identity, reported with every assessment.</summary>
    public string Id { get; }
    /// <summary>Plain-language scope.</summary>
    public string Description { get; }
    /// <summary>Accepted declarations and objects.</summary>
    public IReadOnlyList<string> AcceptedContent { get; }
    /// <summary>Content that makes an input OutOfScope. Being out of scope is not a standards violation.</summary>
    public IReadOnlyList<string> ExcludedContent { get; }

    /// <summary>Every available profile.</summary>
    public static IReadOnlyList<ArchivalInputProfile> All { get; } = Array.AsReadOnly(new[] { SelfGeneratedSimple });
}

/// <summary>Options for <see cref="ArchivalAssessmentService"/>. Values are copied before any input is read.</summary>
public sealed class ArchivalAssessmentOptions
{
    /// <summary>
    /// Existing ZIP/XML/reference/result budgets and local policy, shared with the Preflight services. The assessment always
    /// runs the local preflight with <see cref="ArchivalPreflightOptions.RecognizeKnownCustomTags"/> enabled, whatever this
    /// instance says, so its embedded report carries <see cref="ArchivalPreflightReport.KnownCustomTagsProfileVersion"/>.
    /// </summary>
    public ArchivalPreflightOptions Preflight { get; set; } = new ArchivalPreflightOptions();
    /// <summary>Input profile identity; currently only <see cref="ArchivalInputProfile.SelfGeneratedSimpleId"/>.</summary>
    public string ProfileId { get; set; } = ArchivalInputProfile.SelfGeneratedSimpleId;
    /// <summary>
    /// How a used font without an embedded FontFile is judged for 6.2.6(b)(1)/(b)(2). The default,
    /// <see cref="ArchivalFontExemptionPolicy.ExemptionBasisUnverified"/>, reports Unknown because the edition of
    /// GB/T 33190 that defines font exemptions is not fixed.
    /// </summary>
    public ArchivalFontExemptionPolicy FontExemptionPolicy { get; set; } = ArchivalFontExemptionPolicy.ExemptionBasisUnverified;
    /// <summary>Maximum bytes of one FontFile payload; larger payloads are reported Unknown.</summary>
    public int MaxFontPayloadBytes { get; set; } = 32 * 1024 * 1024;
    /// <summary>Maximum cumulative font payload bytes inspected in one assessment; exceeding it makes the run BudgetExhausted.</summary>
    public long MaxTotalFontPayloadBytes { get; set; } = 128L * 1024 * 1024;

    internal ArchivalFontPolicy SnapshotFonts()
    {
        if (!Enum.IsDefined(typeof(ArchivalFontExemptionPolicy), FontExemptionPolicy) || MaxFontPayloadBytes <= 0 ||
            MaxTotalFontPayloadBytes <= 0) throw new ArgumentOutOfRangeException(nameof(FontExemptionPolicy));
        return new ArchivalFontPolicy(FontExemptionPolicy, MaxFontPayloadBytes, MaxTotalFontPayloadBytes);
    }

    internal ArchivalInputProfile ResolveProfile() =>
        ArchivalInputProfile.All.FirstOrDefault(x => x.Id == ProfileId) ??
        throw new ArgumentException("Unknown archival input profile.", nameof(ProfileId));

    // ArchivalPreflightOptions.Snapshot is internal to Packaging; copy the public values with the same validation.
    internal ArchivalPreflightOptions SnapshotPreflight()
    {
        var source = Preflight ?? throw new ArgumentException("Preflight options are required.", nameof(Preflight));
        var load = source.LoadOptions ?? throw new ArgumentException("LoadOptions is required.", nameof(Preflight));
        if (source.MaxXmlBytes <= 0 || source.MaxSingleXmlBytes <= 0 || source.MaxXmlDepth <= 0 || source.MaxXmlNodes <= 0 ||
            source.MaxReferences <= 0 || source.MaxResults < 12 || load.MaxInputBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(Preflight));
        return new ArchivalPreflightOptions
        {
            RequireSingleDocument = source.RequireSingleDocument, RequireNoEncryption = source.RequireNoEncryption,
            MaxXmlBytes = source.MaxXmlBytes, MaxSingleXmlBytes = source.MaxSingleXmlBytes, MaxXmlDepth = source.MaxXmlDepth,
            MaxXmlNodes = source.MaxXmlNodes, MaxReferences = source.MaxReferences, MaxResults = source.MaxResults,
            RecognizeKnownCustomTags = true,
            LoadOptions = new OfdPackageLoadOptions
            {
                DocumentIndex = load.DocumentIndex, MaxInputBytes = load.MaxInputBytes, MaxEntryCount = load.MaxEntryCount,
                MaxEntryUncompressedBytes = load.MaxEntryUncompressedBytes,
                MaxTotalUncompressedBytes = load.MaxTotalUncompressedBytes, MaxCompressionRatio = load.MaxCompressionRatio,
                MaxPageCount = load.MaxPageCount, MaxAnnotationObjectCount = load.MaxAnnotationObjectCount,
                MaxAnnotationXmlBytes = load.MaxAnnotationXmlBytes
            }
        };
    }
}

/// <summary>How a used font with no embedded FontFile is judged. Host-installed fonts are never evidence.</summary>
public enum ArchivalFontExemptionPolicy
{
    /// <summary>
    /// Default. A missing payload is Unknown with reason <c>FontExemptionBasisUnverified</c>: GB/T 33190's font exemption
    /// and its edition are not fixed, so neither a Fail nor an exemption is asserted.
    /// </summary>
    ExemptionBasisUnverified,
    /// <summary>No exemption is accepted: a used font without a payload is a located Fail on 6.2.6(b)(1) or (b)(2).</summary>
    NoExemptionsAccepted
}

internal sealed class ArchivalFontPolicy
{
    internal ArchivalFontPolicy(ArchivalFontExemptionPolicy exemption, int maxPayloadBytes, long maxTotalBytes)
    { Exemption = exemption; MaxPayloadBytes = maxPayloadBytes; MaxTotalBytes = maxTotalBytes; }
    internal ArchivalFontExemptionPolicy Exemption { get; }
    internal int MaxPayloadBytes { get; }
    internal long MaxTotalBytes { get; }
}
