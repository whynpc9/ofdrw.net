using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

public sealed class OfdDocumentMergeOptions
{
    public bool IncludeTemplates { get; set; } = true;

    public bool IncludeAttachments { get; set; } = true;

    public bool SkipUnsupportedRawElements { get; set; }
}

/// <summary>
/// Merges pages into a self-contained document model. Template content is
/// flattened into each destination page so source package paths and IDs cannot
/// become dangling references.
/// </summary>
public static class OfdDocumentMerger
{
    public static OfdDocumentPackage Merge(
        IEnumerable<OfdDocumentPackage> sources,
        OfdDocumentMergeOptions? options = null)
    {
        if (sources is null)
        {
            throw new ArgumentNullException(nameof(sources));
        }

        options ??= new OfdDocumentMergeOptions();
        var sourceList = sources.ToList();
        if (sourceList.Count == 0)
        {
            throw new ArgumentException("At least one source document is required.", nameof(sources));
        }

        var first = sourceList[0];
        var destination = new OfdDocumentPackage
        {
            Options = CloneOptions(first.Options)
        };
        destination.Options.DocumentId = "Doc_0";

        foreach (var source in sourceList)
        {
            CopyFonts(source, destination);
            if (options.IncludeAttachments)
            {
                CopyAttachments(source, destination);
            }

            foreach (var sourcePage in source.Pages.OrderBy(page => page.Index))
            {
                var page = new OfdPage
                {
                    Index = destination.Pages.Count,
                    XMillimeters = sourcePage.XMillimeters,
                    YMillimeters = sourcePage.YMillimeters,
                    WidthMillimeters = sourcePage.WidthMillimeters,
                    HeightMillimeters = sourcePage.HeightMillimeters
                };

                foreach (var element in EnumerateElements(sourcePage, options.IncludeTemplates))
                {
                    var clone = CloneElement(element, options.SkipUnsupportedRawElements);
                    if (clone is not null)
                    {
                        page.Elements.Add(clone);
                    }
                }

                destination.Pages.Add(page);
            }
        }

        return destination;
    }

    private static IEnumerable<OfdElement> EnumerateElements(OfdPage page, bool includeTemplates)
    {
        if (includeTemplates)
        {
            foreach (var template in page.Templates.Where(template => string.Equals(
                template.ZOrder,
                "Background",
                StringComparison.OrdinalIgnoreCase)))
            {
                foreach (var element in template.Elements)
                {
                    yield return element;
                }
            }
        }

        foreach (var element in page.Elements)
        {
            yield return element;
        }

        if (includeTemplates)
        {
            foreach (var template in page.Templates.Where(template => !string.Equals(
                template.ZOrder,
                "Background",
                StringComparison.OrdinalIgnoreCase)))
            {
                foreach (var element in template.Elements)
                {
                    yield return element;
                }
            }
        }
    }

    private static OfdElement? CloneElement(OfdElement element, bool skipUnsupported)
    {
        if (element is OfdTextElement text)
        {
            var clone = new OfdTextElement
            {
                LayerType = text.LayerType,
                XMillimeters = text.XMillimeters,
                YMillimeters = text.YMillimeters,
                WidthMillimeters = text.WidthMillimeters,
                HeightMillimeters = text.HeightMillimeters,
                Text = text.Text,
                FontName = text.FontName,
                FontSizeMillimeters = text.FontSizeMillimeters,
                Transform = text.Transform?.ToArray(),
                FillColor = CloneColor(text.FillColor)
            };
            foreach (var run in text.Runs)
            {
                clone.Runs.Add(new OfdTextRun
                {
                    Text = run.Text,
                    XMillimeters = run.XMillimeters,
                    YMillimeters = run.YMillimeters,
                    DeltaX = run.DeltaX,
                    DeltaY = run.DeltaY
                });
            }

            return clone;
        }

        if (element is OfdImageElement image)
        {
            return new OfdImageElement
            {
                LayerType = image.LayerType,
                XMillimeters = image.XMillimeters,
                YMillimeters = image.YMillimeters,
                WidthMillimeters = image.WidthMillimeters,
                HeightMillimeters = image.HeightMillimeters,
                FileName = image.FileName,
                MediaType = image.MediaType,
                Data = image.Data.ToArray()
            };
        }

        if (element is OfdPathElement path)
        {
            return new OfdPathElement
            {
                LayerType = path.LayerType,
                XMillimeters = path.XMillimeters,
                YMillimeters = path.YMillimeters,
                WidthMillimeters = path.WidthMillimeters,
                HeightMillimeters = path.HeightMillimeters,
                AbbreviatedData = path.AbbreviatedData,
                Transform = path.Transform?.ToArray(),
                LineWidthMillimeters = path.LineWidthMillimeters,
                Stroke = path.Stroke,
                Fill = path.Fill,
                StrokeColor = CloneColor(path.StrokeColor),
                FillColor = path.FillColor is null ? null : CloneColor(path.FillColor)
            };
        }

        if (skipUnsupported)
        {
            return null;
        }

        throw new NotSupportedException(
            $"Cannot safely merge unsupported OFD page object '{element.GetType().Name}'. " +
            "Set SkipUnsupportedRawElements only when dropping it is acceptable.");
    }

    private static void CopyFonts(OfdDocumentPackage source, OfdDocumentPackage destination)
    {
        foreach (var font in source.Fonts)
        {
            if (destination.Fonts.Any(existing =>
                string.Equals(existing.FontName, font.FontName, StringComparison.OrdinalIgnoreCase) &&
                existing.Bold == font.Bold &&
                existing.Italic == font.Italic))
            {
                continue;
            }

            destination.Fonts.Add(new OfdFontResource
            {
                FontName = font.FontName,
                FamilyName = font.FamilyName,
                Charset = font.Charset,
                Bold = font.Bold,
                Italic = font.Italic,
                FileName = font.FileName,
                Data = font.Data.ToArray()
            });
        }
    }

    private static void CopyAttachments(OfdDocumentPackage source, OfdDocumentPackage destination)
    {
        foreach (var attachment in source.Attachments)
        {
            destination.Attachments.Add(new OfdAttachment
            {
                Name = attachment.Name,
                MediaType = attachment.MediaType,
                Format = attachment.Format,
                CreationDate = attachment.CreationDate,
                ModificationDate = attachment.ModificationDate,
                SizeKilobytes = attachment.SizeKilobytes,
                Visible = attachment.Visible,
                Usage = attachment.Usage,
                IsExternal = attachment.IsExternal,
                ExternalPath = attachment.ExternalPath,
                Data = attachment.Data.ToArray()
            });
        }
    }

    private static OfdDocumentOptions CloneOptions(OfdDocumentOptions source)
    {
        return new OfdDocumentOptions
        {
            DocType = source.DocType,
            Namespace = source.Namespace,
            EnableDeflateCompression = source.EnableDeflateCompression,
            DefaultPageWidthMillimeters = source.DefaultPageWidthMillimeters,
            DefaultPageHeightMillimeters = source.DefaultPageHeightMillimeters,
            Metadata = new OfdMetadata
            {
                Title = source.Metadata.Title,
                Author = source.Metadata.Author,
                Subject = source.Metadata.Subject,
                Keywords = source.Metadata.Keywords,
                Creator = source.Metadata.Creator,
                CreationDate = source.Metadata.CreationDate,
                ModificationDate = source.Metadata.ModificationDate
            }
        };
    }

    private static OfdColor CloneColor(OfdColor color)
    {
        return new OfdColor(color.Red, color.Green, color.Blue, color.Alpha);
    }
}
