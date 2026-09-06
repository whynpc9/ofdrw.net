using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using Ofdrw.Net.Core.Processes;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal sealed class PdfToPpmRasterizer
{
    internal string? LastFailure { get; private set; }
    public async Task<byte[]?> TryRasterizePageAsync(
        string pdfPath, int zeroBasedPageIndex, CancellationToken cancellationToken,
        int dpi = 144, TimeSpan? timeout = null)
    {
        var pageNumber = zeroBasedPageIndex + 1;
        var tempDir = Path.Combine(Path.GetTempPath(), "ofdrw-net-raster", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            var outputPrefix = Path.Combine(tempDir, $"page_{pageNumber}");
            var args = $"-f {pageNumber} -l {pageNumber} -r {Math.Max(72, Math.Min(dpi, 300))} -png -singlefile \"{pdfPath}\" \"{outputPrefix}\"";
            ExternalProcessResult result;
            try
            {
                result = await ExternalProcessRunner.RunAsync(
                    new ProcessStartInfo { FileName = "pdftoppm", Arguments = args },
                    timeout ?? TimeSpan.FromMinutes(2), cancellationToken).ConfigureAwait(false);
            }
            catch (Win32Exception exception)
            {
                LastFailure = exception.Message;
                return null;
            }

            if (result.ExitCode != 0)
            {
                LastFailure = $"exit {result.ExitCode}: {result.Error}";
                return null;
            }

            var imagePath = outputPrefix + ".png";
            if (!File.Exists(imagePath))
            {
                LastFailure = "The renderer produced no page image.";
                return null;
            }

            using var fs = File.OpenRead(imagePath);
            using var ms = new MemoryStream();
            await fs.CopyToAsync(ms, 81920, cancellationToken).ConfigureAwait(false);
            return ms.ToArray();
        }
        finally
        {
            try
            {
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
            catch
            {
                // ignored
            }
        }
    }
}
