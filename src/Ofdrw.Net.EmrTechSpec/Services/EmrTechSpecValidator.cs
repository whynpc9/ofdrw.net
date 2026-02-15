using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Validation;
using Ofdrw.Net.EmrTechSpec.Abstractions;
using Ofdrw.Net.EmrTechSpec.Models;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Validation;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.EmrTechSpec.Services;

public sealed class EmrTechSpecValidator : IEmrTechSpecValidator
{
    private readonly OfdPackageLoader _loader = new();

    public async Task<ValidationReport> ValidateAsync(Stream ofdStream, EmrValidationProfile profile, CancellationToken cancellationToken = default)
    {
        if (ofdStream is null)
        {
            throw new ArgumentNullException(nameof(ofdStream));
        }

        if (profile is null)
        {
            throw new ArgumentNullException(nameof(profile));
        }

        using var buffered = new MemoryStream();
        await ofdStream.CopyToAsync(buffered, 81920, cancellationToken).ConfigureAwait(false);
        buffered.Position = 0;

        var archive = await _loader.LoadAsync(buffered, cancellationToken).ConfigureAwait(false);
        var findings = new List<ValidationFinding>();
        var evidence = new List<ValidationEvidence>
        {
            new ValidationEvidence { Key = "entry_count", Value = archive.EntryNames.Count.ToString() }
        };

        foreach (var issue in OfdPackageStructureChecker.Check(archive))
        {
            findings.Add(new ValidationFinding
            {
                RuleId = issue.Code,
                Clause = "6.1.3",
                Severity = issue.IsError ? ValidationSeverity.Error : ValidationSeverity.Warning,
                Message = issue.Message,
                Location = "package",
                Recommendation = "Fix OFD package layout to match OFD-H single-document structure."
            });
        }

        ValidateDocTypeAndDocRoot(archive, findings);
        ValidateXmlNamespaceAndEncoding(archive, findings);
        await ValidateResourceIntegrityAsync(buffered, findings, evidence, cancellationToken).ConfigureAwait(false);
        AddRecommendations(archive, findings);

        return new ValidationReport
        {
            ProfileVersion = profile.Version,
            Findings = findings,
            Evidence = evidence,
            Passed = findings.All(x => x.Severity != ValidationSeverity.Error)
        };
    }

    private static void ValidateDocTypeAndDocRoot(OfdPackageArchive archive, List<ValidationFinding> findings)
    {
        if (!archive.Contains("OFD.xml"))
        {
            return;
        }

        var xml = XDocument.Parse(archive.ReadUtf8Text("OFD.xml"));
        var root = xml.Root;
        if (root is null)
        {
            findings.Add(new ValidationFinding
            {
                RuleId = "r-doc-root-empty",
                Clause = "6.1.2",
                Severity = ValidationSeverity.Error,
                Message = "OFD root node is missing.",
                Location = "OFD.xml",
                Recommendation = "Create a valid OFD root element."
            });
            return;
        }

        var docType = root.Attribute("DocType")?.Value;
        if (!string.Equals(docType, "OFD-H", StringComparison.Ordinal))
        {
            findings.Add(new ValidationFinding
            {
                RuleId = "r-doc-type-ofd-h",
                Clause = "7-a",
                Severity = ValidationSeverity.Error,
                Message = "DocType must be OFD-H.",
                Location = "OFD.xml@DocType",
                Recommendation = "Set OFD root DocType to OFD-H."
            });
        }

        var ns = root.Name.Namespace;
        var docRoot = root
            .Elements(ns + "DocBody")
            .Elements(ns + "DocRoot")
            .Select(x => x.Value)
            .FirstOrDefault();

        if (!string.Equals(docRoot, "Doc_0/Document.xml", StringComparison.OrdinalIgnoreCase))
        {
            findings.Add(new ValidationFinding
            {
                RuleId = "r-single-doc-doc0",
                Clause = "6.1.2",
                Severity = ValidationSeverity.Error,
                Message = "DocRoot must point to Doc_0/Document.xml.",
                Location = "OFD.xml/DocBody/DocRoot",
                Recommendation = "Use a single document package and point DocRoot to Doc_0/Document.xml."
            });
        }
    }

    private static void ValidateXmlNamespaceAndEncoding(OfdPackageArchive archive, List<ValidationFinding> findings)
    {
        foreach (var entry in archive.EntryNames.Where(x => x.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            var bytes = archive.GetBytes(entry);
            var text = System.Text.Encoding.UTF8.GetString(bytes);
            if (text.Contains("\uFFFD", StringComparison.Ordinal))
            {
                findings.Add(new ValidationFinding
                {
                    RuleId = "r-utf8-readable",
                    Clause = "7-机读信息要求",
                    Severity = ValidationSeverity.Error,
                    Message = "XML content is not valid UTF-8 text.",
                    Location = entry,
                    Recommendation = "Ensure XML files are encoded in UTF-8."
                });
                continue;
            }

            XDocument xml;
            try
            {
                xml = XDocument.Parse(text);
            }
            catch
            {
                findings.Add(new ValidationFinding
                {
                    RuleId = "r-xml-valid",
                    Clause = "6.1.3",
                    Severity = ValidationSeverity.Error,
                    Message = "XML file cannot be parsed.",
                    Location = entry,
                    Recommendation = "Fix XML syntax."
                });
                continue;
            }

            if (!string.Equals(xml.Root?.Name.NamespaceName, "http://www.ofdspec.org", StringComparison.Ordinal))
            {
                findings.Add(new ValidationFinding
                {
                    RuleId = "r-namespace-ofd",
                    Clause = "6.1.3",
                    Severity = ValidationSeverity.Error,
                    Message = "XML namespace must be http://www.ofdspec.org.",
                    Location = entry,
                    Recommendation = "Use the OFD namespace in all XML root elements."
                });
            }
        }
    }

    private static async Task ValidateResourceIntegrityAsync(MemoryStream source, List<ValidationFinding> findings, List<ValidationEvidence> evidence, CancellationToken cancellationToken)
    {
        try
        {
            source.Position = 0;
            var reader = new OfdReader();
            var package = await reader.ReadAsync(source, cancellationToken).ConfigureAwait(false);
            evidence.Add(new ValidationEvidence { Key = "page_count", Value = package.Pages.Count.ToString() });

            for (var i = 0; i < package.Pages.Count; i++)
            {
                var page = package.Pages[i];
                foreach (var image in page.Elements.OfType<Ofdrw.Net.Core.Models.OfdImageElement>())
                {
                    if (image.Data.Length == 0)
                    {
                        findings.Add(new ValidationFinding
                        {
                            RuleId = "r-resource-integrity",
                            Clause = "7-资源文件要求",
                            Severity = ValidationSeverity.Error,
                            Message = "Image resource reference is broken.",
                            Location = $"page[{i}]/{image.ResourceId}",
                            Recommendation = "Ensure every ImageObject points to an existing resource file."
                        });
                    }
                }
            }
        }
        catch (Exception ex)
        {
            findings.Add(new ValidationFinding
            {
                RuleId = "r-parse-document",
                Clause = "5.3.1",
                Severity = ValidationSeverity.Error,
                Message = $"Document cannot be parsed: {ex.Message}",
                Location = "document",
                Recommendation = "Ensure OFD document structure is complete and consistent."
            });
        }
    }

    private static void AddRecommendations(OfdPackageArchive archive, List<ValidationFinding> findings)
    {
        if (!archive.Contains("Doc_0/Signatures.xml") && !archive.EntryNames.Any(x => x.Contains("/Signs/", StringComparison.OrdinalIgnoreCase)))
        {
            findings.Add(new ValidationFinding
            {
                RuleId = "w-signature-recommended",
                Clause = "5.3.1/8",
                Severity = ValidationSeverity.Warning,
                Message = "No signature artifacts found. Effective medical record files are recommended to include signatures.",
                Location = "Doc_0/Signs",
                Recommendation = "Add digital signature/signature metadata when document enters effective state."
            });
        }

        if (!archive.Contains("Encryptions.xml"))
        {
            findings.Add(new ValidationFinding
            {
                RuleId = "w-encryption-optional",
                Clause = "8",
                Severity = ValidationSeverity.Warning,
                Message = "Encryptions.xml not found. This is acceptable unless encryption is required by scene policy.",
                Location = "Encryptions.xml",
                Recommendation = "Use Encryptions.xml for temporary protected OFD scenarios."
            });
        }

        if (!archive.EntryNames.Any(x => x.Contains("DocInfo_", StringComparison.OrdinalIgnoreCase)))
        {
            findings.Add(new ValidationFinding
            {
                RuleId = "w-version-history",
                Clause = "7-版本要求",
                Severity = ValidationSeverity.Warning,
                Message = "No DocInfo_N version metadata found.",
                Location = "Doc_0",
                Recommendation = "Keep historical version metadata or attach historical versions as internal attachments."
            });
        }
    }
}
