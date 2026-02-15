using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Validation;

public static class OfdPackageStructureChecker
{
    public static IReadOnlyList<OfdPackageStructureIssue> Check(OfdPackageArchive archive)
    {
        var issues = new List<OfdPackageStructureIssue>();

        if (!archive.Contains(OfdConstants.OfdRootFile))
        {
            issues.Add(new OfdPackageStructureIssue
            {
                Code = "missing_ofd_xml",
                Message = "OFD.xml is required.",
                IsError = true
            });
        }

        if (!archive.Contains("Doc_0/Document.xml"))
        {
            issues.Add(new OfdPackageStructureIssue
            {
                Code = "missing_doc0_document",
                Message = "Doc_0/Document.xml is required.",
                IsError = true
            });
        }

        var docs = archive.EntryNames
            .Select(GetTopDirectory)
            .Where(x => x.StartsWith("Doc_", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (docs.Any(x => !x.Equals(OfdConstants.DefaultDocId, StringComparison.OrdinalIgnoreCase)))
        {
            issues.Add(new OfdPackageStructureIssue
            {
                Code = "multiple_docs_detected",
                Message = "Only single document Doc_0 is allowed.",
                IsError = true
            });
        }

        var pageEntries = archive.FindByPrefix("Doc_0/Pages").ToList();
        if (!pageEntries.Any())
        {
            issues.Add(new OfdPackageStructureIssue
            {
                Code = "missing_pages",
                Message = "Doc_0/Pages is required.",
                IsError = true
            });
        }

        return issues;
    }

    private static string GetTopDirectory(string path)
    {
        var idx = path.IndexOf('/');
        return idx <= 0 ? path : path.Substring(0, idx);
    }
}
