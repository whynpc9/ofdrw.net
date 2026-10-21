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
| SkiaSharp | 3.119.1 | MIT | https://go.microsoft.com/fwlink/?linkid=868515 |
| HarfBuzzSharp | 14.2.1.301 | MIT | https://go.microsoft.com/fwlink/?linkid=868515 |
| HarfBuzzSharp.NativeAssets.macOS | 14.2.1.301 | MIT | https://go.microsoft.com/fwlink/?linkid=868515 |
| HarfBuzzSharp.NativeAssets.Win32 | 14.2.1.301 | MIT | https://go.microsoft.com/fwlink/?linkid=868515 |
| HarfBuzzSharp.NativeAssets.Linux | 14.2.1.301 | MIT | https://go.microsoft.com/fwlink/?linkid=868515 |
| BouncyCastle.Cryptography | 2.6.2 | MIT | https://www.bouncycastle.org/stable/nuget/csharp/website |
| SkiaSharp.NativeAssets.Linux.NoDependencies (19 tests/sample only) | 3.119.1 | MIT, bundled native notices | https://github.com/mono/SkiaSharp/tree/v3.119.1 |
| SixLabors.Fonts | 1.0.1 | Apache-2.0 | https://github.com/SixLabors/Fonts |
| SixLabors.ImageSharp | 2.1.13 | Apache-2.0 | https://github.com/SixLabors/ImageSharp |

The ImageSharp row describes **the pinned 2.1.13 package**: its nuspec declares
`Apache-2.0`, and the [v2.1.13 source LICENSE](https://github.com/SixLabors/ImageSharp/blob/v2.1.13/LICENSE)
contains Apache License 2.0. Current ImageSharp mainline documentation describes
a Six Labors Split License; any ImageSharp version update therefore needs a new
license review rather than carrying this row forward automatically.

## Optional SES dev/interop extension

`Ofdrw.Net.Signatures.SesInterop` alone adds `BouncyCastle.Cryptography` at the
exact version `2.6.2` (MIT), source https://github.com/bcgit/bc-csharp.
Core, Writer, Converter and the default CLI do not depend on this extension.
The actual NuGet metadata and MIT license are retained under
`docs/evidence/ses-interop/`; the extension package also carries the license.

Transitive dependencies are resolved by NuGet and may change as direct
dependencies are updated. The reviewed release inventory in
`docs/third-party-dependency-baseline.json` records the resolved package
versions and package-declared license metadata. A changed dependency graph
requires renewed license review before release; older packages with only a
license URL are not automatically assigned an SPDX expression.

DOCX conversion invokes a separately installed LibreOffice executable. Ofdrw.Net
does not bundle or redistribute LibreOffice; consumers are responsible for its
installation and for reviewing the applicable LibreOffice license notices.

On macOS, DOCX conversion may instead automate a separately installed Microsoft
Word application for higher layout fidelity. Ofdrw.Net does not bundle Microsoft
Word or Microsoft Office fonts. The LibreOffice backend may reference installed
Office fonts in place, but does not copy them into the package or repository.


The Issue 05 development/E2E fixtures include unmodified LXGW WenKai and Noto
Sans fonts under SIL OFL 1.1, with pinned upstream commit URLs and SHA-256 in
`e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/upstream-manifest.json`. Their
original license files accompany the fonts. These fixtures are not included
in the production NuGet packages. Modified LXGW subsets replace Reserved Font
Names while retaining copyright and license records. Subsetting does not grant
embedding rights. The managed subset implementation introduces no third-party
runtime backend for the existing Layout and Converter packages. HarfBuzz and
fonttools were previously used only for independent development probes; the
new optional `Ofdrw.Net.Text.HarfBuzz` adapter is a separate runtime dependency.
It pins the managed package and macOS, Win32 and Linux native asset packages
to the same version. The native packages carry `THIRD-PARTY-NOTICES.txt`,
including the HarfBuzz Old MIT notice. This does not grant font embedding rights.

The compiled font subset bidi profile derives factual scalar ranges from the
Unicode 17.0.0 Character Database (DerivedBidiClass.txt), licensed under
Unicode-3.0. The source URL/SHA-256 are recorded in UnicodeFontSubsetProfile.cs;
`docs/licenses/UNICODE-LICENSE.txt` accompanies every NuGet artifact.


The optional producer-event adapter has its own SkiaSharp dependency. Default
Layout/Converter products do not depend on it. Skia native binaries include
additional third-party notices shipped by their NativeAssets packages; see
`THIRD-PARTY-NOTICES.txt` in the resolved native package. NativeAssets
macOS/Win32 are SkiaSharp transitive runtime dependencies. No custom Skia build
or native bridge is redistributed by this adapter.

## Optional password envelope extension

`Ofdrw.Net.Crypto.Password` directly references `BouncyCastle.Cryptography`
2.6.2, MIT, source https://github.com/bcgit/bc-csharp/tree/release-2.6.2.
Its nuspec pins upstream commit `b4f2f6ad76bcd1f11f365ee50cc7447fbce79077`.
This dependency is not added to Core, Writer or the default Converter metapackage.

## Native process supervisor

The first-party `Ofdrw.ProcessSupervisor` Linux x64 and ARM64 executables
statically link musl 1.2.5 (MIT), source
https://musl.libc.org/releases/musl-1.2.5.tar.gz, SHA-256
`a9a118bbe84d8764da0ea0d28b3ab3fae8477fc7e4085d90102b8596fc7c75e4`.
The complete upstream copyright/license notice and applicable individual
source notices are retained in `docs/licenses/MUSL-COPYRIGHT.txt` and in the
Core package. The pinned GCC build tools are not redistributed; explicit
`-nostdlib` linking and retained link maps establish that no libgcc or other
compiler support library is statically included. macOS executables link only
to operating-system libSystem/libproc; those libraries are not redistributed.
The exact source, toolchain digests, flags, headers and binary hashes are in
`native/Ofdrw.ProcessSupervisor/assets/manifest.json`.
