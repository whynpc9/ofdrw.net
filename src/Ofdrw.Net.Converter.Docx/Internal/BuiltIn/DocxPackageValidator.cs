using System;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

internal static class DocxPackageValidator
{
    internal static void Validate(string inputPath, DocxConversionOptions options)
    {
        var length = new FileInfo(inputPath).Length;
        if (length <= 0)
        {
            throw new InvalidDataException("The DOCX input is empty.");
        }

        if (length > options.MaxInputBytes)
        {
            throw new InvalidDataException(
                $"The DOCX input exceeds the configured {options.MaxInputBytes} byte limit.");
        }

        try
        {
            using var archive = ZipFile.OpenRead(inputPath);
            if (archive.Entries.Count > options.MaxPackagePartCount)
            {
                throw new InvalidDataException(
                    $"The DOCX package contains more than {options.MaxPackagePartCount} entries.");
            }

            long expandedBytes = 0;
            foreach (var entry in archive.Entries)
            {
                ValidateEntryName(entry.FullName);
                expandedBytes = checked(expandedBytes + entry.Length);
                if (expandedBytes > options.MaxExpandedBytes)
                {
                    throw new InvalidDataException(
                        $"The DOCX package expands beyond {options.MaxExpandedBytes} bytes.");
                }

                if (entry.CompressedLength > 0 && entry.Length > 1024 * 1024 &&
                    entry.Length / entry.CompressedLength > 200)
                {
                    throw new InvalidDataException(
                        $"The DOCX package entry '{entry.FullName}' has an unsafe compression ratio.");
                }

                if (entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    ValidateRelationships(entry);
                }
            }

            if (archive.GetEntry("[Content_Types].xml") is null ||
                archive.GetEntry("word/document.xml") is null)
            {
                throw new InvalidDataException("The input is not a WordprocessingML DOCX package.");
            }
        }
        catch (InvalidDataException)
        {
            throw;
        }
        catch (Exception ex) when (
            ex is IOException or UnauthorizedAccessException or XmlException or OverflowException)
        {
            throw new InvalidDataException("The DOCX package could not be validated safely.", ex);
        }
    }

    private static void ValidateEntryName(string name)
    {
        var normalized = name.Replace('\\', '/');
        if (normalized.StartsWith("/", StringComparison.Ordinal) ||
            normalized.IndexOf("../", StringComparison.Ordinal) >= 0 ||
            normalized.Equals("..", StringComparison.Ordinal))
        {
            throw new InvalidDataException("The DOCX package contains an unsafe entry path.");
        }
    }

    private static void ValidateRelationships(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 4 * 1024 * 1024
        });

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element &&
                reader.LocalName == "Relationship" &&
                string.Equals(reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
            {
                var relationshipType = reader.GetAttribute("Type");
                if (relationshipType?.EndsWith("/hyperlink", StringComparison.OrdinalIgnoreCase) == true)
                {
                    continue;
                }

                throw new InvalidDataException(
                    "External DOCX relationships are not allowed by managed DOCX processing.");
            }
        }
    }
}
