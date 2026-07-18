#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="${1:-0.1.0-preview.3}"
OUT_DIR="$ROOT_DIR/artifacts/nuget"
export DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1

echo "[E2E] Root: $ROOT_DIR"
echo "[E2E] Package version: $VERSION"

mkdir -p "$OUT_DIR"

pack_project() {
  local project="$1"
  dotnet pack "$project" -c Release -o "$OUT_DIR" \
    -p:Version="$VERSION" \
    -p:PackageVersion="$VERSION" \
    -p:RestoreSources="https://api.nuget.org/v3/index.json" \
    --disable-build-servers \
    -m:1 \
    /nodeReuse:false \
    /p:UseSharedCompilation=false \
    --nologo
}

pack_project "$ROOT_DIR/src/Ofdrw.Net.Core/Ofdrw.Net.Core.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Packaging/Ofdrw.Net.Packaging.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Reader/Ofdrw.Net.Reader.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Converter.Abstractions/Ofdrw.Net.Converter.Abstractions.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Converter.Pdf/Ofdrw.Net.Converter.Pdf.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Converter.Svg/Ofdrw.Net.Converter.Svg.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Signatures/Ofdrw.Net.Signatures.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Converter/Ofdrw.Net.Converter.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Cli/Ofdrw.Net.Cli.csproj"

echo "[E2E] Local packages are ready in $OUT_DIR"

OFDRW_REPO_ROOT="$ROOT_DIR" dotnet run \
  --project "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Pdf.E2E/Ofdrw.Net.Converter.Pdf.E2E.csproj" \
  -c Release \
  --disable-build-servers \
  --configfile "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Pdf.E2E/NuGet.config" \
  -m:1 \
  /nodeReuse:false \
  /p:UseSharedCompilation=false \
  -p:OfdrwPackageVersion="$VERSION"
