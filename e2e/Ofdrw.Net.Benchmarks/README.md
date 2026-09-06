# Conversion benchmark probe

This small runner compares **installed package versions** using identical inputs.
It is a local diagnostic benchmark, not a statistically controlled cross-machine ranking.

- `generate <directory>` creates fixed 1/10/100-page PDFs (432 x 576 points).
- `pdf <input.pdf> <report.json> <trials> <label>` measures PDF -> OFD at 150 DPI.
- `docx <input.docx> <report.json> <trials> <label>` measures BuiltIn DualLayer DOCX -> OFD at 150 DPI.

Build copies of the project in separate directories with separate `NUGET_PACKAGES`
caches and `-p:OfdrwPackageVersion=<version>`. Use a NuGet configuration that maps
`Ofdrw.Net.*` exclusively to the candidate/baseline artifact feeds. Both runs must
use the same generated input files. Avoid running the two variants concurrently.

Each process warms the relevant engine before the timed trials, records managed
allocations, output size, source hash, loaded library version, and process peak
working set. Native allocations are reflected in working set, not the managed
allocation counter. Save and compare actual output page images as well as timing.

Apply the repository's writable `DOTNET_CLI_HOME`, telemetry opt-out and serial
build instructions. Use separate build and `dotnet run --no-build --no-restore`
stages; do not pass MSBuild flags through application arguments.
