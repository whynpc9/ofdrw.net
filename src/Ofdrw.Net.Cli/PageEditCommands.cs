using System.Globalization;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

internal static partial class Cli
{
    // ofdrw delete-pages <input> <output> --pages 2,4-5 [--document N]
    private static async Task<int> DeletePagesAsync(CliOptions options, CancellationToken token)
    {
        var pages = (ParsePages(options.Pages) ?? throw new ArgumentException("--pages is required for delete-pages."))
            .Distinct().ToArray();
        var removed = 0;
        await EditPackageAsync(options.InputPath!, options.OutputPath!, options.DocumentIndex ?? 0, package =>
        {
            if (pages.Any(index => index >= package.Pages.Count))
                throw new ArgumentOutOfRangeException(nameof(options.Pages), "A selected page does not exist in the document.");
            if (pages.Length >= package.Pages.Count) throw new ArgumentException("delete-pages cannot remove every page.");
            OfdDocumentEditor.RemovePages(package, pages);
            removed = pages.Length;
        }, token).ConfigureAwait(false);
        WriteSuccessNotification($"Deleted {removed} page(s) {options.InputPath} -> {options.OutputPath}");
        return 0;
    }

    // ofdrw crop <input> <output> --box X,Y,W,H [--pages 1,3] [--document N]
    private static async Task<int> CropAsync(CliOptions options, CancellationToken token)
    {
        var box = ParseBox(options.Box ?? throw new ArgumentException("--box is required for crop."));
        var selected = ParsePages(options.Pages)?.Distinct().ToArray();
        var cropped = 0;
        await EditPackageAsync(options.InputPath!, options.OutputPath!, options.DocumentIndex ?? 0, package =>
        {
            var ordered = package.Pages.OrderBy(page => page.Index).ToList();
            var targets = selected ?? Enumerable.Range(0, ordered.Count).ToArray();
            if (targets.Any(index => index >= ordered.Count))
                throw new ArgumentOutOfRangeException(nameof(options.Pages), "A selected page does not exist in the document.");
            foreach (var index in targets) OfdDocumentEditor.CropPage(ordered[index], box[0], box[1], box[2], box[3]);
            cropped = targets.Length;
        }, token).ConfigureAwait(false);
        WriteSuccessNotification($"Cropped {cropped} page(s) {options.InputPath} -> {options.OutputPath}");
        return 0;
    }

    private static double[] ParseBox(string value)
    {
        var parts = value.Split(',');
        if (parts.Length != 4) throw new ArgumentException("--box must be X,Y,W,H in millimeters.");
        var box = parts.Select(part => double.TryParse(part.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var number) &&
            !double.IsNaN(number) && !double.IsInfinity(number) ? number : throw new ArgumentException("--box values must be finite numbers.")).ToArray();
        if (box[2] <= 0 || box[3] <= 0) throw new ArgumentException("--box width and height must be positive.");
        return box;
    }

    // Same read → edit → atomic write shape as reorder; the input closes before commit, so in-place edits work.
    private static async Task EditPackageAsync(string inputPath, string outputPath, int documentIndex,
        Action<OfdDocumentPackage> edit, CancellationToken token)
    {
        if (!File.Exists(inputPath)) throw new FileNotFoundException("Input file does not exist.", inputPath);
        OfdDocumentPackage package;
        await using (var input = File.OpenRead(inputPath))
            package = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = documentIndex }, token).ConfigureAwait(false);
        edit(package);
        await using var output = new AtomicOutput(outputPath);
        var result = await new OfdPackageWriter().WriteWithResultAsync(package, output.Stream, token).ConfigureAwait(false);
        foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine($"Warning: {diagnostic}");
        await output.CommitAsync(token).ConfigureAwait(false);
    }
}
