using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Validation;

public static class OfdPackageStructureChecker
{
    public static IReadOnlyList<OfdPackageStructureIssue> Check(OfdPackageArchive archive)
    {
        if (archive is null)
        {
            throw new ArgumentNullException(nameof(archive));
        }

        var issues = new List<OfdPackageStructureIssue>();

        if (!archive.Contains(OfdConstants.OfdRootFile))
        {
            issues.Add(new OfdPackageStructureIssue
            {
                Code = "missing_ofd_xml",
                Message = "OFD.xml is required.",
                IsError = true
            });
            return issues;
        }

        XDocument ofdXml;
        try
        {
            ofdXml = XDocument.Parse(archive.ReadUtf8Text(OfdConstants.OfdRootFile));
        }
        catch (Exception ex)
        {
            issues.Add(new OfdPackageStructureIssue
            {
                Code = "invalid_ofd_xml",
                Message = $"OFD.xml is not valid XML: {ex.Message}",
                IsError = true
            });
            return issues;
        }

        var docRoots = ofdXml.Descendants()
            .Where(x => x.Name.LocalName == "DocRoot")
            .Select(x => Normalize(x.Value))
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        if (docRoots.Count == 0)
        {
            issues.Add(new OfdPackageStructureIssue
            {
                Code = "missing_doc_root",
                Message = "OFD.xml must contain at least one DocBody/DocRoot entry.",
                IsError = true
            });
            return issues;
        }

        foreach (var docRoot in docRoots)
        {
            if (!archive.Contains(docRoot))
            {
                issues.Add(new OfdPackageStructureIssue
                {
                    Code = "missing_document_xml",
                    Message = $"Referenced document entry is missing: {docRoot}",
                    IsError = true
                });
                continue;
            }

            ValidateDocument(archive, docRoot, issues);
        }

        return issues;
    }

    private static void ValidateDocument(
        OfdPackageArchive archive,
        string docRoot,
        ICollection<OfdPackageStructureIssue> issues)
    {
        XDocument documentXml;
        try
        {
            documentXml = XDocument.Parse(archive.ReadUtf8Text(docRoot));
        }
        catch (Exception ex)
        {
            issues.Add(new OfdPackageStructureIssue
            {
                Code = "invalid_document_xml",
                Message = $"{docRoot} is not valid XML: {ex.Message}",
                IsError = true
            });
            return;
        }

        var pages = documentXml.Descendants()
            .Where(x => x.Name.LocalName == "Page" && x.Attribute("BaseLoc") is not null)
            .ToList();
        if (pages.Count == 0)
        {
            issues.Add(new OfdPackageStructureIssue
            {
                Code = "missing_pages",
                Message = $"{docRoot} does not reference any page content.",
                IsError = true
            });
        }

        foreach (var page in pages)
        {
            var pagePath = Resolve(docRoot, page.Attribute("BaseLoc")!.Value);
            if (!archive.Contains(pagePath))
            {
                issues.Add(new OfdPackageStructureIssue
                {
                    Code = "missing_page_content",
                    Message = $"Referenced page entry is missing: {pagePath}",
                    IsError = true
                });
            }
        }

        foreach (var referenceName in new[] { "PublicRes", "DocumentRes", "Annotations", "CustomTags", "Attachments", "Extensions" })
        {
            foreach (var reference in documentXml.Descendants().Where(x => x.Name.LocalName == referenceName))
            {
                var path = Resolve(docRoot, reference.Value);
                if (!string.IsNullOrWhiteSpace(reference.Value) && !archive.Contains(path))
                {
                    issues.Add(new OfdPackageStructureIssue
                    {
                        Code = "missing_referenced_entry",
                        Message = $"{referenceName} references a missing entry: {path}",
                        IsError = true
                    });
                }
            }
        }
    }

    private static string Resolve(string basePath, string relativePath)
    {
        if (relativePath.StartsWith("/", StringComparison.Ordinal))
        {
            return Normalize(relativePath);
        }

        var directoryIndex = basePath.LastIndexOf('/');
        var directory = directoryIndex < 0 ? string.Empty : basePath.Substring(0, directoryIndex);
        var segments = new Stack<string>();
        foreach (var segment in $"{directory}/{relativePath}".Split('/'))
        {
            if (string.IsNullOrWhiteSpace(segment) || segment == ".")
            {
                continue;
            }

            if (segment == "..")
            {
                if (segments.Count > 0)
                {
                    segments.Pop();
                }

                continue;
            }

            segments.Push(segment);
        }

        return string.Join("/", segments.Reverse());
    }

    private static string Normalize(string path)
    {
        return path.Replace('\\', '/').TrimStart('/');
    }
}
