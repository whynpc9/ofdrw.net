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
    /// <summary>Existing ZIP/XML/reference/result budgets and local policy, shared with the Preflight services.</summary>
    public ArchivalPreflightOptions Preflight { get; set; } = new ArchivalPreflightOptions();
    /// <summary>Input profile identity; currently only <see cref="ArchivalInputProfile.SelfGeneratedSimpleId"/>.</summary>
    public string ProfileId { get; set; } = ArchivalInputProfile.SelfGeneratedSimpleId;

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
