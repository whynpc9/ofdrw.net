#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="${1:-0.1.0-preview.5}"
if [[ $# -gt 0 ]]; then shift; fi
if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$ ]]; then
  echo "Expected a SemVer package version." >&2
  exit 1
fi
OUT_DIR="$ROOT_DIR/artifacts/package-e2e/$VERSION/packages"
RESULT_DIR="$ROOT_DIR/artifacts/package-e2e/$VERSION/output"
CONSUME_ONLY=false
while [[ $# -gt 0 ]]; do
  case "$1" in
    --consume-only) CONSUME_ONLY=true; shift ;;
    --packages-dir) OUT_DIR="$2"; shift 2 ;;
    --output-dir) RESULT_DIR="$2"; shift 2 ;;
    *) echo "Unknown package E2E option: $1" >&2; exit 1 ;;
  esac
done
mkdir -p "$OUT_DIR" "$RESULT_DIR"
OUT_DIR="$(cd "$OUT_DIR" && pwd)"
RESULT_DIR="$(cd "$RESULT_DIR" && pwd)"
TASK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/ofdrw-package-e2e.XXXXXX")"
trap 'rm -rf "$TASK_DIR"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$TASK_DIR/dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1
export DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
SOURCE_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
BUILD_FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --nologo)

if [[ "$CONSUME_ONLY" == false ]]; then
  dotnet restore "$ROOT_DIR/Ofdrw.Net.sln" "${BUILD_FLAGS[@]}"
  dotnet build "$ROOT_DIR/Ofdrw.Net.sln" -c Release --no-restore \
    -p:Version="$VERSION" -p:PackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
  dotnet pack "$ROOT_DIR/Ofdrw.Net.sln" -c Release --no-build --no-restore -o "$OUT_DIR" \
    -p:Version="$VERSION" -p:PackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
fi
python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$OUT_DIR" "$VERSION"

# A clean consumer directory and package cache prevent old same-version DLLs or
# project references from satisfying the test. Only this artifact feed may supply
# Ofdrw.Net packages; third-party dependencies may reuse the existing local feed.
mkdir -p "$TASK_DIR/consumer"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs" "$TASK_DIR/consumer/Program.cs"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Pdf.E2E/Ofdrw.Net.Converter.Pdf.E2E.csproj" "$TASK_DIR/consumer/Consumer.csproj"
python3 - "$TASK_DIR/NuGet.Config" "$OUT_DIR" "$SOURCE_CACHE" <<'PY'
import sys
import xml.etree.ElementTree as ET
root = ET.Element('configuration')
sources = ET.SubElement(root, 'packageSources')
ET.SubElement(sources, 'clear')
for key, value in [('local-artifacts', sys.argv[2]), ('cached-third-party', sys.argv[3]), ('nuget.org', 'https://api.nuget.org/v3/index.json')]:
    ET.SubElement(sources, 'add', key=key, value=value)
mapping = ET.SubElement(root, 'packageSourceMapping')
for key, pattern in [('local-artifacts', 'Ofdrw.Net.*'), ('cached-third-party', '*'), ('nuget.org', '*')]:
    source = ET.SubElement(mapping, 'packageSource', key=key)
    ET.SubElement(source, 'package', pattern=pattern)
ET.ElementTree(root).write(sys.argv[1], encoding='utf-8', xml_declaration=True)
PY
export NUGET_PACKAGES="$TASK_DIR/packages"
CONSUMER="$TASK_DIR/consumer/Consumer.csproj"
dotnet restore "$CONSUMER" --configfile "$TASK_DIR/NuGet.Config" -p:OfdrwPackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
dotnet build "$CONSUMER" -c Release --no-restore -p:OfdrwPackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
dotnet tool install Ofdrw.Net.Cli --tool-path "$TASK_DIR/tools" --version "$VERSION" --configfile "$TASK_DIR/NuGet.Config"
"$TASK_DIR/tools/ofdrw" docx-to-ofd \
  "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx" "$RESULT_DIR/cli-native.ofd"
OFDRW_REPO_ROOT="$ROOT_DIR" OFDRW_E2E_OUTPUT_DIR="$RESULT_DIR" dotnet run \
  --project "$CONSUMER" -c Release --no-build --no-restore
python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$OUT_DIR" "$VERSION" --verify-manifest
cp "$OUT_DIR/package-manifest.json" "$RESULT_DIR/package-manifest.json"
echo "[E2E] Verified package bytes and output: $RESULT_DIR"
