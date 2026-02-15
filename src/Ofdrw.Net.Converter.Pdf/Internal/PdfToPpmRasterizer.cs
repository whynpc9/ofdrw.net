using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal sealed class PdfToPpmRasterizer
{
    public async Task<byte[]?> TryRasterizePageAsync(string pdfPath, int zeroBasedPageIndex, CancellationToken cancellationToken)
    {
        var pageNumber = zeroBasedPageIndex + 1;
        var tempDir = Path.Combine(Path.GetTempPath(), "ofdrw-net-raster", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            var outputPrefix = Path.Combine(tempDir, $"page_{pageNumber}");
            var args = $"-f {pageNumber} -l {pageNumber} -png -singlefile \"{pdfPath}\" \"{outputPrefix}\"";

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "pdftoppm",
                    Arguments = args,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                }
            };

            try
            {
                process.Start();
            }
            catch
            {
                return null;
            }

            using (process)
            {
                await Task.Run(() => process.WaitForExit(), cancellationToken).ConfigureAwait(false);
                if (process.ExitCode != 0)
                {
                    return null;
                }
            }

            var imagePath = outputPrefix + ".png";
            if (!File.Exists(imagePath))
            {
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
