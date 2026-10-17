# Deterministic visual corpus

`Program.cs` in the package consumer creates four synthetic, privacy-free OFDs:
ticket, invoice, template, and a valid edge layout with touching/clipped shapes.
Each declares Chinese and English text with Noto Sans CJK SC; resolved
fonts remain specific to the reviewed platform baseline. The invoice includes
stroked table borders and proportional body text.
Each is exported **OFD → PDF → PNG at 144 DPI**. All four page-1 PNGs are
compared with reviewed files in `golden/<platform>/` using normalized ImageMagick RMSE
with a maximum of `0.035` for the full page and separately for title/body
regions. The consumer fails if any golden is absent or any region differs.
Generated OFDs, PDFs, PNGs, diff images and CSV metrics remain under the E2E
output directory.

The macOS golden images were created with the .NET 10.0.203 SDK and
`pdftoppm` 26.09.0. Linux golden images are captured from the Ubuntu CI
package E2E with its pinned test font. Separate baselines avoid accepting a
large cross-platform font drift as a rendering regression. Compare environment
upgrades separately and review each golden diff before replacing it. This checks the four named samples only. It is
an automated regression gate and does not count as macOS Preview acceptance.

To refresh after an intentional rendering change, run the package E2E, inspect
the freshly generated PDFs in Preview, review the generated PNGs and diffs,
copy only approved images from `artifacts/package-e2e/<version>/output/visual-corpus/`
to the appropriate platform folder, then rerun the E2E without modifications. Retain the acceptance
record using `docs/preview-acceptance-template.md`.

Windows uses an explicitly embedded, SHA-256 checked Noto face from the existing
trusted-font fixture, with explicit character advances measured from that same
face. The first capture lacked DeltaX: PDF export inferred spacing, but direct
Suwell OFD rendering overlapped characters, so those images were not accepted.
The Windows DOCX package fixture also receives the trusted Noto directory and
disables Office private-font discovery; an isolated font install is not a system
font install. Historical Linux/macOS fixture configuration is unchanged.
Run the same package consumer with native Windows tools:

```powershell
python scripts/run-converter-package-e2e.py 0.1.0-preview.8
```

To consume an existing immutable eleven-package feed, add `--consume-only
--packages-dir <feed>`. CI consumes the exact packages produced by the Linux job.
The runner restores source dependency metadata for license checks, then builds an
isolated PackageReference consumer with a fresh package cache; it does not build
product DLLs in consume-only mode. Logs, hashes and `package-e2e-run.json` are
saved under the output directory.

For a candidate baseline, add `--visual-corpus-only --capture-visual-baseline`.
This renders all four pages but reports `CAPTURE_ONLY_NOT_VALIDATED`; it never
copies images into `golden/`. Inspect every PNG and freshly open every PDF in
Preview, then copy only approved images and rerun without capture mode. The normal
gate still fails for a missing baseline or a difference above the threshold.
`--visual-corpus-only` without capture runs the existing golden comparisons.

Set `OFDRW_PDFTOPPM`, `OFDRW_CONVERT`, `OFDRW_IDENTIFY`, `OFDRW_COMPARE`, and
`OFDRW_RSVG_CONVERT` to absolute executable paths for isolated tool installations.
On Windows, ImageMagick commands otherwise use `magick`, avoiding the unrelated
Windows `convert.exe`. The CI tool versions and bootstrap hash are pinned in
`.github/workflows/ci.yml`.
