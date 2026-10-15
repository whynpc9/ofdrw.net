#!/usr/bin/env bash
set -euo pipefail
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUTPUT_DIR="${1:-$ROOT_DIR/artifacts/graphics/current}"
FONT_DIR="${OFDRW_GRAPHICS_FONTS:-$ROOT_DIR/artifacts/graphics-fonts}"
mkdir -p "$OUTPUT_DIR"
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$(mktemp -d "${TMPDIR:-/tmp}/ofd-graphics-dotnet.XXXXXX")}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false)
dotnet build "$ROOT_DIR/e2e/Ofdrw.Net.Graphics.Sample" -c Release "${FLAGS[@]}"
dotnet build "$ROOT_DIR/e2e/Ofdrw.Net.Graphics.E2E" -c Release "${FLAGS[@]}"
dotnet "$ROOT_DIR/e2e/Ofdrw.Net.Graphics.Sample/bin/Release/net10.0/Ofdrw.Net.Graphics.Sample.dll" "$OUTPUT_DIR/graphics.ofd" "$FONT_DIR/Ofdrw-CI-NotoSansCJKsc-Regular.ttf"
dotnet "$ROOT_DIR/e2e/Ofdrw.Net.Graphics.E2E/bin/Release/net10.0/Ofdrw.Net.Graphics.E2E.dll" "$ROOT_DIR" "$OUTPUT_DIR" "$FONT_DIR"
python3 "$ROOT_DIR/scripts/verify-graphics-evidence.py" "$OUTPUT_DIR" "$ROOT_DIR"
