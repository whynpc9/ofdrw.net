# R00 ordinary PageBlock loss reproduction

Source baseline: `6aa8183ea31e193f500fb9d9717677b15e86c5d3` on macOS, .NET SDK 10.0.401. This fixture is deliberately one 100 × 100 mm page: a top-level white PathObject and an ordinary `PageBlock` containing a 50 × 50 mm blue PathObject at (10, 10). No fonts, images, templates or annotations are involved. [`Program.cs`](Program.cs) creates the OFD, reads it again, and runs the PDF, standalone SVG and HTML exporters. The generated files stay under ignored `artifacts/`.

Run from the repository root after setting the `DOTNET_CLI_HOME`, first-time and telemetry variables required by `AGENTS.md`:

```sh
dotnet build docs/evidence/r00-pageblock/repro.csproj -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet docs/evidence/r00-pageblock/bin/Release/net10.0/repro.dll artifacts/r00/pageblock-review
pdftoppm -f 1 -singlefile -png -r 96 artifacts/r00/pageblock-review/pageblock.pdf artifacts/r00/pageblock-review/pageblock
```

Observed after build and fresh execution on 2026-10-21: Reader returned `OfdPathElement,OfdRawElement`; standalone SVG completed but did not contain the blue fill; PDF completed with one blank white page, inspected in the page 1 PNG and in macOS Preview after opening the **newly generated** PDF. The Preview window was closed after inspection. HTML rejected `OfdRawElement` with `NotSupportedException` before public output and left its output stream at zero bytes. PNG/JPEG export through PDF would inherit this loss; this fixture did not separately invoke that image API. The current PDF/SVG behavior is a **confirmed visual defect for this sample**, so this is not an overall visual acceptance. R01 must repair or explicitly fail these exports before claiming support for ordinary PageBlock.

| Generated item | SHA-256 from this run |
| --- | --- |
| `artifacts/r00/pageblock-review/pageblock.ofd` | `1eb1a83113937e133722de6e0f2fdfe92b60c955db2d1df39bbb9c9fc92470df` |
| `artifacts/r00/pageblock-review/pageblock.pdf` | `187104684c2cb3eb82b233e080a03f5abdf252beebc40bda19f5d971728f4783` |
| `artifacts/r00/pageblock-review/pageblock.svg` | `466a017000694defcf9e6ea9ba8f14f6cd41658ee96a31f969bfd881ddd11c3f` |
| `artifacts/r00/pageblock-review/pageblock.png` | `60f858846dae466f12e53e9dc16218b413ef1653558ae0587cd81c6f11b81621` |

The OFD ZIP may have a different hash on another run because the writer timestamps its entries. The semantic fixture and the blue rectangle expectation are fixed in source. Do not archive a second copy of the generated package or all build outputs in Git.
