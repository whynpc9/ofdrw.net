# Third-Party Notices

Ofdrw.Net depends on the following third-party packages. This notice records
their package-declared licenses and source locations; it does not select a
license for Ofdrw.Net itself.

| Package | Version | License | Source |
| --- | ---: | --- | --- |
| PdfSharpCore | 1.3.67 | MIT | https://github.com/ststeiger/PdfSharpCore |
| PdfPig | 0.1.15 | Apache-2.0 | https://github.com/UglyToad/PdfPig |
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
