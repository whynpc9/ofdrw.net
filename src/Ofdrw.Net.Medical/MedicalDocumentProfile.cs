using System;
using System.Linq;
using Ofdrw.Net.Archival;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Medical;

/// <summary>
/// Names of the <c>DocInfo/CustomDatas/CustomData</c> fields that carry the GB/T 48666 7.3(a)/(b) items the standard
/// does not map to a GB/T 33190 element. The standard defines no field names; <see cref="Default"/> is this library's
/// convention, not a standard schema. Use the institution's names when they exist.
/// </summary>
public sealed class MedicalMetadataMapping
{
    /// <summary>This library's convention: <c>EMR.PatientID</c>, <c>EMR.FileType</c>, <c>EMR.SecretLevel</c>.</summary>
    public static MedicalMetadataMapping Default => new MedicalMetadataMapping
    { PatientField = "EMR.PatientID", FileTypeField = "EMR.FileType", SecretLevelField = "EMR.SecretLevel" };

    /// <summary>CustomData name of the patient (service object) identifier. Null leaves 7.3(a) patient Unknown.</summary>
    public string? PatientField { get; set; }
    /// <summary>CustomData name of the file type. Null leaves 7.3(b) file type Unknown.</summary>
    public string? FileTypeField { get; set; }
    /// <summary>CustomData name of the secrecy level. Null leaves 7.3(b) secrecy level Unknown.</summary>
    public string? SecretLevelField { get; set; }

    internal MedicalMetadataMapping Snapshot() => new MedicalMetadataMapping
    { PatientField = PatientField, FileTypeField = FileTypeField, SecretLevelField = SecretLevelField };
}

/// <summary>Metadata required by GB/T 48666 7.3(a) and 7.3(b).</summary>
public sealed class MedicalDocumentMetadata
{
    /// <summary>Unique document identifier, written as <c>DocInfo/DocID</c>. Required.</summary>
    public string DocumentIdentifier { get; set; } = string.Empty;
    /// <summary>Patient (service object) identifier. Required.</summary>
    public string PatientId { get; set; } = string.Empty;
    /// <summary>Responsible party, written as <c>DocInfo/Author</c>. Required.</summary>
    public string ResponsibleParty { get; set; } = string.Empty;
    /// <summary>Creation time, written as <c>DocInfo/CreationDate</c>. Required.</summary>
    public DateTimeOffset CreationTime { get; set; }
    /// <summary>Content keywords, written as <c>DocInfo/Keywords</c>. Required.</summary>
    public string Keywords { get; set; } = string.Empty;
    /// <summary>File type, for example the record category. Required.</summary>
    public string FileType { get; set; } = string.Empty;
    /// <summary>Secrecy level. Required.</summary>
    public string SecretLevel { get; set; } = string.Empty;
    /// <summary>Optional title.</summary>
    public string? Title { get; set; }
}

/// <summary>
/// Opt-in writer profile for GB/T 48666 OFD-H output. It only configures <see cref="OfdDocumentOptions"/>: the caller
/// still generates and writes the package, and should assess the written bytes with <see cref="MedicalAssessmentService"/>.
/// It does not embed fonts, sign, add attachments or create double-layer pages.
/// </summary>
public static class MedicalDocumentProfile
{
    /// <summary>
    /// Sets <c>DocType=OFD-H</c> (7.1(a)), the <c>/2016</c> namespace (6.2), the <c>Doc_0</c> folder (6.1.2), Deflate
    /// compression (6.1.3), linearized entry order (6.1.3) and the 7.3(a)/(b) metadata.
    /// </summary>
    /// <param name="options">Options to change in place, typically <c>FlowDocument.Options</c> or a new instance.</param>
    /// <param name="metadata">Required medical metadata.</param>
    /// <param name="mapping">CustomData names; null uses <see cref="MedicalMetadataMapping.Default"/>. Every field must be named.</param>
    /// <returns>The same <paramref name="options"/> instance.</returns>
    /// <exception cref="ArgumentException">A required value or mapping name is empty, or a mapped CustomData already holds another value.</exception>
    public static OfdDocumentOptions Apply(OfdDocumentOptions options, MedicalDocumentMetadata metadata, MedicalMetadataMapping? mapping = null)
    {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (metadata == null) throw new ArgumentNullException(nameof(metadata));
        var names = (mapping ?? MedicalMetadataMapping.Default).Snapshot();
        Require(metadata.DocumentIdentifier, nameof(metadata.DocumentIdentifier));
        Require(metadata.PatientId, nameof(metadata.PatientId));
        Require(metadata.ResponsibleParty, nameof(metadata.ResponsibleParty));
        Require(metadata.Keywords, nameof(metadata.Keywords));
        Require(metadata.FileType, nameof(metadata.FileType));
        Require(metadata.SecretLevel, nameof(metadata.SecretLevel));
        if (metadata.CreationTime == default) throw new ArgumentException("CreationTime is required.", nameof(metadata));
        Require(names.PatientField, nameof(names.PatientField));
        Require(names.FileTypeField, nameof(names.FileTypeField));
        Require(names.SecretLevelField, nameof(names.SecretLevelField));
        if (new[] { names.PatientField, names.FileTypeField, names.SecretLevelField }.Distinct(StringComparer.Ordinal).Count() != 3)
            throw new ArgumentException("Mapped CustomData names must be distinct.", nameof(mapping));

        var custom = options.Metadata.CustomDatas ?? new OfdCustomDatas();
        var fields = new[] { (names.PatientField!, metadata.PatientId), (names.FileTypeField!, metadata.FileType), (names.SecretLevelField!, metadata.SecretLevel) };
        // Check every mapped name before inserting any, so a rejected apply leaves the caller's CustomDatas untouched.
        foreach (var (name, value) in fields)
            if (custom.Items.Any(x => x.Name == name && x.Value != value))
                throw new ArgumentException("CustomData '" + name + "' already holds a different value.", nameof(options));
        foreach (var (name, value) in fields)
            if (!custom.Items.Any(x => x.Name == name)) custom.Add(name, value);

        options.DocType = MedicalDocTypes.Medical;
        options.Namespace = OfdConstants.StandardNamespace;
        options.DocumentId = OfdConstants.DefaultDocId;
        options.EnableDeflateCompression = true;
        options.LinearizedEntryOrder = true;
        options.Metadata.DocId = metadata.DocumentIdentifier;
        options.Metadata.Author = metadata.ResponsibleParty;
        options.Metadata.CreationDate = metadata.CreationTime;
        options.Metadata.Keywords = metadata.Keywords;
        if (metadata.Title != null) options.Metadata.Title = metadata.Title;
        options.Metadata.CustomDatas = custom;
        return options;
    }

    private static void Require(string? value, string name)
    {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException(name + " is required.", name);
    }
}

/// <summary>Options for <see cref="MedicalAssessmentService"/>.</summary>
public sealed class MedicalAssessmentOptions
{
    /// <summary>Options of the embedded GB/T 42133 assessment; its Preflight budgets also bound this package's reads.</summary>
    public ArchivalAssessmentOptions Archival { get; set; } = new ArchivalAssessmentOptions();
    /// <summary>CustomData names for 7.3(a)/(b); null leaves those items Unknown.</summary>
    public MedicalMetadataMapping? Metadata { get; set; } = MedicalMetadataMapping.Default;
    /// <summary>Institution limit for embedded associated objects (7.2(e), 7.4(c)); the standard suggests 512 MB.</summary>
    public long MaxEmbeddedObjectBytes { get; set; } = 512L * 1024 * 1024;
}
