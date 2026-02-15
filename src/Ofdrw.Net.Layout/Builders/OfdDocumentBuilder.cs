using System;
using System.Linq;
using Ofdrw.Net.Core.Interfaces;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Builders;

public sealed class OfdDocumentBuilder : IOfdDocumentBuilder
{
    private readonly OfdDocumentPackage _package = new();

    public OfdDocumentOptions Options => _package.Options;

    public IOfdDocumentBuilder SetOptions(OfdDocumentOptions options)
    {
        _package.Options = options ?? throw new ArgumentNullException(nameof(options));
        return this;
    }

    public IOfdDocumentBuilder AddPage(OfdPage page)
    {
        if (page is null)
        {
            throw new ArgumentNullException(nameof(page));
        }

        if (page.WidthMillimeters <= 0 || page.HeightMillimeters <= 0)
        {
            throw new ArgumentException("Page width/height must be greater than 0.", nameof(page));
        }

        if (page.Index < 0)
        {
            page.Index = _package.Pages.Count;
        }

        _package.Pages.Add(page);
        return this;
    }

    public IOfdDocumentBuilder AddText(int pageIndex, OfdTextElement element)
    {
        if (element is null)
        {
            throw new ArgumentNullException(nameof(element));
        }

        var page = GetOrCreatePage(pageIndex);
        page.Elements.Add(element);
        return this;
    }

    public IOfdDocumentBuilder AddImage(int pageIndex, OfdImageElement element)
    {
        if (element is null)
        {
            throw new ArgumentNullException(nameof(element));
        }

        var page = GetOrCreatePage(pageIndex);
        page.Elements.Add(element);
        return this;
    }

    public IOfdDocumentBuilder AddAttachment(OfdAttachment attachment)
    {
        if (attachment is null)
        {
            throw new ArgumentNullException(nameof(attachment));
        }

        _package.Attachments.Add(attachment);
        return this;
    }

    public OfdDocumentPackage Build()
    {
        var normalized = new OfdDocumentPackage
        {
            Options = _package.Options
        };

        foreach (var page in _package.Pages.OrderBy(p => p.Index))
        {
            normalized.Pages.Add(page);
        }

        foreach (var attachment in _package.Attachments)
        {
            normalized.Attachments.Add(attachment);
        }

        foreach (var pair in _package.CustomTags)
        {
            normalized.CustomTags[pair.Key] = pair.Value;
        }

        return normalized;
    }

    private OfdPage GetOrCreatePage(int pageIndex)
    {
        if (pageIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(pageIndex), "Page index must be non-negative.");
        }

        var existing = _package.Pages.FirstOrDefault(x => x.Index == pageIndex);
        if (existing is not null)
        {
            return existing;
        }

        var page = new OfdPage
        {
            Index = pageIndex,
            WidthMillimeters = Options.DefaultPageWidthMillimeters,
            HeightMillimeters = Options.DefaultPageHeightMillimeters
        };
        _package.Pages.Add(page);
        return page;
    }
}
