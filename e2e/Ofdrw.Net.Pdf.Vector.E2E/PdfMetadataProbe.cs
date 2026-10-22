using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

internal static class PdfMetadataProbe
{
    internal static async Task Run(string output, PdfFixture fixture, string firstPage, string secondPage)
    {
        var source = fixture.Create(new[] { firstPage, secondPage },
            catalogEntries: "/Outlines 13 0 R /Names << /EmbeddedFiles 16 0 R >>",
            pageEntriesByIndex: new[] { "/Annots [19 0 R]", "" },
            trailingObjects: new[]
            {
                "<< /Type /Outlines /First 14 0 R /Last 14 0 R /Count 2 >>",
                "<< /Title (Chapter A) /Parent 13 0 R /First 15 0 R /Last 15 0 R /Count 1 /Dest [11 0 R /Fit] >>",
                "<< /Title (Child) /Parent 14 0 R /Dest [9 0 R /XYZ 30 500 null] >>",
                "<< /Names [(same.bin) 17 0 R] >>",
                "<< /Type /Filespec /F (same.bin) /EF << /F 18 0 R >> >>",
                "<< /Type /EmbeddedFile /Length 6 >>\nstream\nGLOBAL\nendstream",
                "<< /Type /Annot /Subtype /FileAttachment /Rect [340 205 365 230] /Contents (page note) /FS 20 0 R >>",
                "<< /Type /Filespec /F (same.bin) /EF << /F 21 0 R >> >>",
                "<< /Type /EmbeddedFile /Length 4 >>\nstream\nPAGE\nendstream"
            });
        File.WriteAllBytes(Path.Combine(output, "metadata-source.pdf"), source);
        var reports = new List<object>();
        foreach (var mode in new[] { "dual", "vector" })
        {
            using var input = new MemoryStream(source);
            using var ofd = new MemoryStream();
            var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
            if (mode == "vector")
            {
                var result = await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options })
                    .ConvertWithResultAsync(input, ofd, new[] { 1, 0 });
                if (result.Pages.Any(page => !page.IsNative)) throw new InvalidDataException("R07 vector metadata page was not native.");
            }
            else await new PdfToOfdConverter(options).ConvertAsync(input, ofd, new[] { 1, 0 });
            File.WriteAllBytes(Path.Combine(output, "metadata-" + mode + ".ofd"), ofd.ToArray());
            ofd.Position = 0;
            var package = await new OfdReader().ReadAsync(ofd);
            if (package.Pages.Count != 2 || package.Outlines?.Items.Count != 1 ||
                !ReferenceEquals(((OfdGoToAction)package.Outlines.Items[0].Actions!.Items.Single()).Destination!.Page, package.Pages[0]) ||
                !ReferenceEquals(((OfdGoToAction)package.Outlines.Items[0].Children.Single().Actions!.Items.Single()).Destination!.Page, package.Pages[1]) ||
                package.Attachments.Count != 2 || package.Attachments[0].Name != "same.bin" || package.Attachments[1].Name != "same.bin (2)" ||
                !package.Attachments[0].Data.SequenceEqual("GLOBAL"u8.ToArray()) || !package.Attachments[1].Data.SequenceEqual("PAGE"u8.ToArray()) ||
                package.Pages[0].Annotations is not null || package.Pages[1].Annotations?.Items.SingleOrDefault()?.AttachmentId != package.Attachments[1].Id)
                throw new InvalidDataException("R07 package readback failed.");
            ofd.Position = 0;
            await using (var pdf = File.Create(Path.Combine(output, "metadata-" + mode + ".pdf")))
                await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            reports.Add(new { mode, sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(source)),
                ofdSha256 = Convert.ToHexStringLower(SHA256.HashData(ofd.ToArray())),
                attachments = package.Attachments.Select(file => new { file.Name, sha256 = Convert.ToHexStringLower(SHA256.HashData(file.Data)) }) });
        }
        File.WriteAllText(Path.Combine(output, "metadata-report.json"), JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }
}
