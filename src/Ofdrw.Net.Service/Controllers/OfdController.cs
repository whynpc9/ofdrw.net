using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.EmrTechSpec.Abstractions;
using Ofdrw.Net.EmrTechSpec.Services;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Service.Contracts;

namespace Ofdrw.Net.Service.Controllers;

[ApiController]
[Route("api/v1/ofd")]
public sealed class OfdController : ControllerBase
{
    private readonly IPdfToOfdConverter _pdfToOfdConverter;
    private readonly IOfdToPdfConverter _ofdToPdfConverter;
    private readonly IEmrTechSpecValidator _validator;
    private readonly EmrValidationProfileRepository _profileRepository;

    public OfdController(
        IPdfToOfdConverter pdfToOfdConverter,
        IOfdToPdfConverter ofdToPdfConverter,
        IEmrTechSpecValidator validator,
        EmrValidationProfileRepository profileRepository)
    {
        _pdfToOfdConverter = pdfToOfdConverter;
        _ofdToPdfConverter = ofdToPdfConverter;
        _validator = validator;
        _profileRepository = profileRepository;
    }

    [HttpPost("generate")]
    public async Task<IActionResult> GenerateAsync([FromBody] GenerateOfdRequest request, CancellationToken cancellationToken)
    {
        var builder = new OfdDocumentBuilder();
        if (request.Options is not null)
        {
            builder.SetOptions(request.Options);
        }

        for (var pageIndex = 0; pageIndex < request.Pages.Count; pageIndex++)
        {
            var sourcePage = request.Pages[pageIndex];
            var page = new OfdPage
            {
                Index = pageIndex,
                WidthMillimeters = sourcePage.WidthMillimeters > 0 ? sourcePage.WidthMillimeters : builder.Options.DefaultPageWidthMillimeters,
                HeightMillimeters = sourcePage.HeightMillimeters > 0 ? sourcePage.HeightMillimeters : builder.Options.DefaultPageHeightMillimeters
            };

            foreach (var text in sourcePage.Texts)
            {
                page.Elements.Add(new OfdTextElement
                {
                    Text = text.Text,
                    FontName = text.FontName,
                    FontSizeMillimeters = text.FontSizeMillimeters,
                    XMillimeters = text.XMillimeters,
                    YMillimeters = text.YMillimeters,
                    WidthMillimeters = text.WidthMillimeters,
                    HeightMillimeters = text.HeightMillimeters
                });
            }

            foreach (var image in sourcePage.Images)
            {
                if (string.IsNullOrWhiteSpace(image.Base64Data))
                {
                    continue;
                }

                page.Elements.Add(new OfdImageElement
                {
                    XMillimeters = image.XMillimeters,
                    YMillimeters = image.YMillimeters,
                    WidthMillimeters = image.WidthMillimeters,
                    HeightMillimeters = image.HeightMillimeters,
                    MediaType = string.IsNullOrWhiteSpace(image.MediaType) ? "image/png" : image.MediaType,
                    FileName = image.FileName,
                    Data = Convert.FromBase64String(image.Base64Data)
                });
            }

            builder.AddPage(page);
        }

        foreach (var attachment in request.Attachments)
        {
            builder.AddAttachment(new OfdAttachment
            {
                Name = attachment.Name,
                MediaType = attachment.MediaType,
                IsExternal = attachment.IsExternal,
                ExternalPath = attachment.ExternalPath,
                Data = attachment.IsExternal || string.IsNullOrWhiteSpace(attachment.Base64Data)
                    ? []
                    : Convert.FromBase64String(attachment.Base64Data)
            });
        }

        var package = builder.Build();
        foreach (var tag in request.CustomTags)
        {
            package.CustomTags[tag.Key] = tag.Value;
        }

        await using var output = new MemoryStream();
        var writer = new OfdPackageWriter();
        await writer.WriteAsync(package, output, cancellationToken).ConfigureAwait(false);
        output.Position = 0;
        return File(output.ToArray(), "application/ofd", "generated.ofd");
    }

    [HttpPost("convert/pdf-to-ofd")]
    [Consumes("multipart/form-data")]
    public async Task<IActionResult> PdfToOfdAsync([FromForm] IFormFile pdf, [FromForm] string? pages, CancellationToken cancellationToken)
    {
        if (pdf is null || pdf.Length == 0)
        {
            return BadRequest("pdf file is required");
        }

        var pageList = ParsePages(pages);

        await using var input = pdf.OpenReadStream();
        await using var output = new MemoryStream();
        await _pdfToOfdConverter.ConvertAsync(input, output, pageList, cancellationToken).ConfigureAwait(false);

        return File(output.ToArray(), "application/ofd", Path.ChangeExtension(pdf.FileName, ".ofd"));
    }

    [HttpPost("convert/ofd-to-pdf")]
    [Consumes("multipart/form-data")]
    public async Task<IActionResult> OfdToPdfAsync([FromForm] IFormFile ofd, [FromForm] string? pages, CancellationToken cancellationToken)
    {
        if (ofd is null || ofd.Length == 0)
        {
            return BadRequest("ofd file is required");
        }

        var pageList = ParsePages(pages);

        await using var input = ofd.OpenReadStream();
        await using var output = new MemoryStream();
        await _ofdToPdfConverter.ConvertAsync(input, output, pageList, cancellationToken).ConfigureAwait(false);

        return File(output.ToArray(), "application/pdf", Path.ChangeExtension(ofd.FileName, ".pdf"));
    }

    [HttpPost("validate/emr-tech-spec")]
    [Consumes("multipart/form-data")]
    public async Task<IActionResult> ValidateEmrTechSpecAsync([FromForm] IFormFile ofd, CancellationToken cancellationToken)
    {
        if (ofd is null || ofd.Length == 0)
        {
            return BadRequest("ofd file is required");
        }

        var profile = _profileRepository.GetDefaultProfile();
        await using var input = ofd.OpenReadStream();
        var report = await _validator.ValidateAsync(input, profile, cancellationToken).ConfigureAwait(false);
        return Ok(report);
    }

    [HttpGet("validate/profiles/emr-ofd-h-202x")]
    public IActionResult GetProfile()
    {
        var profile = _profileRepository.GetDefaultProfile();
        return Ok(profile);
    }

    private static IReadOnlyList<int>? ParsePages(string? pages)
    {
        if (string.IsNullOrWhiteSpace(pages))
        {
            return null;
        }

        var values = pages.Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(x => int.TryParse(x, out var page) ? (int?)page : null)
            .Where(x => x.HasValue)
            .Select(x => x!.Value)
            .ToList();

        return values.Count == 0 ? null : values;
    }
}
