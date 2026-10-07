#!/usr/bin/env bash
set -euo pipefail
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUTPUT_DIR="${1:-$ROOT_DIR/artifacts/ses-interop/e2e}"
VERSION="${2:-0.1.0-ses-interop-test}"
if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$ ]]; then
  echo "Expected a SemVer package version." >&2; exit 1
fi
mkdir -p "$OUTPUT_DIR"
OUTPUT_DIR="$(cd "$OUTPUT_DIR" && pwd)"
TASK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/ofd-ses-consumer.XXXXXX")"
trap 'rm -rf "$TASK_DIR"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$TASK_DIR/dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
SOURCE_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
export NUGET_PACKAGES="$SOURCE_CACHE"
FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --nologo)
mkdir -p "$OUTPUT_DIR/packages/sdk" "$OUTPUT_DIR/packages/optional" "$TASK_DIR/consumer"
bash "$ROOT_DIR/scripts/run-converter-package-e2e.sh" "$VERSION" --packages-dir "$OUTPUT_DIR/packages/sdk" --output-dir "$OUTPUT_DIR/sdk-output"
dotnet pack "$ROOT_DIR/src/Ofdrw.Net.Signatures.SesInterop" -c Release -o "$OUTPUT_DIR/packages/optional" -p:Version="$VERSION" -p:PackageVersion="$VERSION" "${FLAGS[@]}"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Signatures.SesInterop.E2E/Program.cs" "$TASK_DIR/consumer/Program.cs"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Signatures.SesInterop.E2E/Ofdrw.Net.Signatures.SesInterop.E2E.csproj" "$TASK_DIR/consumer/Consumer.csproj"
python3 - "$TASK_DIR/NuGet.Config" "$OUTPUT_DIR/packages/sdk" "$OUTPUT_DIR/packages/optional" "$SOURCE_CACHE" <<'PY'
import sys,xml.etree.ElementTree as E
root=E.Element('configuration'); sources=E.SubElement(root,'packageSources'); E.SubElement(sources,'clear')
mapping=E.SubElement(root,'packageSourceMapping')
for key,path,pattern in [('sdk',sys.argv[2],'Ofdrw.Net.*'),('ses',sys.argv[3],'Ofdrw.Net.Signatures.SesInterop'),('cache',sys.argv[4],'*'),('nuget','https://api.nuget.org/v3/index.json','*')]:
 E.SubElement(sources,'add',key=key,value=path)
 E.SubElement(E.SubElement(mapping,'packageSource',key=key),'package',pattern=pattern)
E.ElementTree(root).write(sys.argv[1],encoding='utf-8',xml_declaration=True)
PY
export NUGET_PACKAGES="$TASK_DIR/packages"
dotnet restore "$TASK_DIR/consumer/Consumer.csproj" --configfile "$TASK_DIR/NuGet.Config" -p:OfdrwSesPackageVersion="$VERSION" "${FLAGS[@]}"
dotnet build "$TASK_DIR/consumer/Consumer.csproj" -c Release --no-restore -p:OfdrwSesPackageVersion="$VERSION" "${FLAGS[@]}"
dotnet tool install Ofdrw.Net.Cli --tool-path "$TASK_DIR/tools" --version "$VERSION" --configfile "$TASK_DIR/NuGet.Config"
FONT_DIR="${OFDRW_SES_FONTS:-$ROOT_DIR/artifacts/ses-interop/fonts}"
dotnet "$TASK_DIR/consumer/bin/Release/net10.0/Consumer.dll" "$ROOT_DIR" "$OUTPUT_DIR/files" "$FONT_DIR"
for mode in native default; do
  for version in v1 v4; do
    name="$mode-$version"
    set +e
    "$TASK_DIR/tools/ofdrw" verify-signatures "$OUTPUT_DIR/files/$name.ofd" > "$OUTPUT_DIR/files/$name-cli.txt" 2>&1
    result=$?
    set -e
    if [[ "$result" -ne 2 ]]; then cat "$OUTPUT_DIR/files/$name-cli.txt"; echo "Expected CLI reference-only exit 2, got $result" >&2; exit 1; fi
  done
done
python3 "$ROOT_DIR/scripts/verify-ses-interop-evidence.py" "$OUTPUT_DIR" "$VERSION" "$TASK_DIR/consumer/obj/project.assets.json" "$TASK_DIR/tools/.store/ofdrw.net.cli/$VERSION/ofdrw.net.cli/$VERSION/tools/net10.0/any/Ofdrw.Net.Cli.deps.json"
python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$OUTPUT_DIR/packages/sdk" "$VERSION" --verify-manifest
