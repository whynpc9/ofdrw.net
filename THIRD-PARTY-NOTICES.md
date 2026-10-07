# Third-Party Notices

Ofdrw.Net depends on the following third-party packages. This notice records
their package-declared licenses and source locations; it does not select a
license for Ofdrw.Net itself.

| Package | Version | License | Source |
| --- | ---: | --- | --- |
| Docnet.Core | 2.6.0 | MIT | https://github.com/GowenGit/docnet |
| DocumentFormat.OpenXml | 3.5.1 | MIT | https://github.com/dotnet/Open-XML-SDK |
| MigraDocCore.Rendering | 1.3.67 | MIT | https://github.com/ststeiger/PdfSharpCore |
| PdfSharpCore | 1.3.67 | MIT | https://github.com/ststeiger/PdfSharpCore |
| PdfPig | 0.1.15 | Apache-2.0 | https://github.com/UglyToad/PdfPig |
| SkiaSharp (optional 19) | 3.119.1 | MIT | https://github.com/mono/SkiaSharp/tree/v3.119.1 |
| SkiaSharp.NativeAssets.Linux.NoDependencies (19 tests/sample only) | 3.119.1 | MIT, bundled native notices | https://github.com/mono/SkiaSharp/tree/v3.119.1 |
| SixLabors.Fonts | 1.0.1 | Apache-2.0 | https://github.com/SixLabors/Fonts |
| SixLabors.ImageSharp | 2.1.13 | Apache-2.0 | https://github.com/SixLabors/ImageSharp |

Transitive dependencies are resolved by NuGet and may change as direct
dependencies are updated. Consumers should use the generated dependency graph
and the corresponding package metadata when performing a release compliance
review.

DOCX conversion invokes a separately installed LibreOffice executable. Ofdrw.Net
does not bundle or redistribute LibreOffice; consumers are responsible for its
installation and for reviewing the applicable LibreOffice license notices.

On macOS, DOCX conversion may instead automate a separately installed Microsoft
Word application for higher layout fidelity. Ofdrw.Net does not bundle Microsoft
Word or Microsoft Office fonts. The LibreOffice backend may reference installed
Office fonts in place, but does not copy them into the package or repository.

The optional producer-event adapter has its own SkiaSharp dependency. Default
Layout/Converter products do not depend on it. Skia native binaries include
additional third-party notices shipped by their NativeAssets packages; see
`THIRD-PARTY-NOTICES.txt` in the resolved native package. NativeAssets
macOS/Win32 are SkiaSharp transitive runtime dependencies. No custom Skia build
or native bridge is redistributed by this adapter.
