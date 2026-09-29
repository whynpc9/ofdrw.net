# Deterministic visual corpus

`Program.cs` in the package consumer creates four synthetic, privacy-free OFDs:
ticket, invoice, template, and a valid edge layout with touching/clipped shapes.
Each is exported **OFD → PDF → PNG at 144 DPI**. All four page-1 PNGs are
compared with reviewed files in `golden/` using normalized ImageMagick RMSE
with a maximum of `0.035` for the full page and separately for title/body
regions. The consumer fails if any golden is absent or any region differs.
Generated OFDs, PDFs, PNGs, diff images and CSV metrics remain under the E2E
output directory.

The golden images were created on macOS with the .NET 10.0.203 SDK,
`pdftoppm` 26.09.0, and the package E2E. CI installs its pinned test font and
uses `ubuntu-latest`; compare environment upgrades separately and review each
golden diff before replacing it. This checks the four named samples only. It is
an automated regression gate and does not count as macOS Preview acceptance.

To refresh after an intentional rendering change, run the package E2E, inspect
the freshly generated PDFs in Preview, review the generated PNGs and diffs,
copy only approved images from `artifacts/package-e2e/<version>/output/visual-corpus/`
to `golden/`, then rerun the E2E without modifications. Retain the acceptance
record using `docs/preview-acceptance-template.md`.
