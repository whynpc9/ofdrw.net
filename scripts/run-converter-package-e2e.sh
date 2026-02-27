#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="${1:-0.1.0-preview.1}"
OUT_DIR="$ROOT_DIR/artifacts/nuget"

echo "[E2E] Root: $ROOT_DIR"
echo "[E2E] Package version: $VERSION"

mkdir -p "$OUT_DIR"

pack_project() {
  local project="$1"
  dotnet pack "$project" -c Release -o "$OUT_DIR" \
    -p:Version="$VERSION" \
    -p:PackageVersion="$VERSION" \
    --nologo
}

pack_project "$ROOT_DIR/src/Ofdrw.Net.Core/Ofdrw.Net.Core.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Packaging/Ofdrw.Net.Packaging.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Reader/Ofdrw.Net.Reader.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Converter.Abstractions/Ofdrw.Net.Converter.Abstractions.csproj"
pack_project "$ROOT_DIR/src/Ofdrw.Net.Converter.Pdf/Ofdrw.Net.Converter.Pdf.csproj"

echo "[E2E] Local packages are ready in $OUT_DIR"

OFDRW_REPO_ROOT="$ROOT_DIR" dotnet run \
  --project "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Pdf.E2E/Ofdrw.Net.Converter.Pdf.E2E.csproj" \
  -c Release \
  --configfile "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Pdf.E2E/NuGet.config" \
  -p:OfdrwPackageVersion="$VERSION"
