#!/usr/bin/env bash
set -euo pipefail
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUTPUT_DIR="${1:-$ROOT_DIR/docs/evidence/document-tools/files}"
mkdir -p "$OUTPUT_DIR"
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$(mktemp -d "${TMPDIR:-/tmp}/ofd-tools-dotnet.XXXXXX")}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
export DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false)
dotnet build "$ROOT_DIR/e2e/Ofdrw.Net.DocumentTools.E2E" "${FLAGS[@]}"
dotnet build "$ROOT_DIR/src/Ofdrw.Net.Cli" "${FLAGS[@]}"
FONT_DIR="${OFDRW_DOCUMENT_TOOLS_FONTS:-$ROOT_DIR/artifacts/document-tools-fonts}"
dotnet "$ROOT_DIR/e2e/Ofdrw.Net.DocumentTools.E2E/bin/Debug/net10.0/Ofdrw.Net.DocumentTools.E2E.dll" "$ROOT_DIR" "$OUTPUT_DIR" "$FONT_DIR"
SAMPLE_FONT="$(python3 - "$OUTPUT_DIR/signed.ofd" <<'PYFONT'
import sys,zipfile,xml.etree.ElementTree as E
with zipfile.ZipFile(sys.argv[1]) as z:
 root=E.fromstring(z.read('Doc_0/PublicRes.xml'))
 print(next(n for n in root.iter() if n.tag.rsplit('}',1)[-1]=='Font').attrib['FontName'])
PYFONT
)"
CLI="$ROOT_DIR/src/Ofdrw.Net.Cli/bin/Debug/net10.0/Ofdrw.Net.Cli.dll"
printf 'existing-output' > "$OUTPUT_DIR/cli-custom-tags-protected-output.bin"
if dotnet "$CLI" mix "$OUTPUT_DIR/cli-custom-tags-protected-output.bin" "$OUTPUT_DIR/custom-tags-input.ofd" 1 "$OUTPUT_DIR/custom-tags-conflict.ofd" 1 > "$OUTPUT_DIR/cli-custom-tags.rejection.txt" 2>&1; then
  echo "Expected CustomTags conflict rejection." >&2
  exit 1
fi
python3 - "$OUTPUT_DIR/cli-custom-tags-protected-output.bin" "$OUTPUT_DIR/cli-custom-tags.rejection.txt" <<'PYTAGS'
import pathlib,sys
assert pathlib.Path(sys.argv[1]).read_bytes()==b'existing-output'
assert 'CustomTags' in pathlib.Path(sys.argv[2]).read_text()
PYTAGS
dotnet "$CLI" watermark "$OUTPUT_DIR/signed.ofd" "$OUTPUT_DIR/cli-text.ofd" --pages 1 --text 'DRAFT 草稿' --font "$SAMPLE_FONT" --x 60 --y 250 --width 100 --height 16 --font-size 7 --layer watermark --layer-type Foreground
LAYER_ID="$(python3 - "$OUTPUT_DIR/cli-text.ofd" <<'PY'
import sys,zipfile,xml.etree.ElementTree as E
with zipfile.ZipFile(sys.argv[1]) as z:
 root=E.fromstring(z.read('Doc_0/Pages/Page_0/Content.xml'))
 print([n for n in root.iter() if n.tag.rsplit('}',1)[-1]=='Layer'][-1].attrib['ID'])
PY
)"
dotnet "$CLI" watermark "$OUTPUT_DIR/cli-text.ofd" "$OUTPUT_DIR/cli-watermark.ofd" --pages 1 --image "$ROOT_DIR/e2e/Ofdrw.Net.DocumentTools.E2E/mark.png" --x 175 --y 215 --width 15 --height 15 --layer "$LAYER_ID" --layer-type Foreground
dotnet "$CLI" split "$OUTPUT_DIR/signed.ofd" "$OUTPUT_DIR/cli-split.ofd" --pages 2,1
dotnet "$CLI" mix "$OUTPUT_DIR/cli-mix.ofd" "$OUTPUT_DIR/signed.ofd" 1 "$OUTPUT_DIR/overlay.ofd" 1
dotnet "$CLI" mix "$OUTPUT_DIR/cli-mix-custom-tags.ofd" "$OUTPUT_DIR/custom-tags-input.ofd" 1 "$OUTPUT_DIR/overlay.ofd" 1
dotnet "$CLI" clean-signatures "$OUTPUT_DIR/signed.ofd" "$OUTPUT_DIR/cli-clean.ofd"
dotnet "$CLI" verify-signatures "$OUTPUT_DIR/cli-clean.ofd" > "$OUTPUT_DIR/cli-clean-verification.txt"
dotnet "$CLI" merge "$OUTPUT_DIR/cli-merged.ofd" "$OUTPUT_DIR/cli-watermark.ofd" "$OUTPUT_DIR/overlay.ofd"
for name in cli-watermark cli-split cli-mix cli-mix-custom-tags cli-clean cli-merged; do
  dotnet "$CLI" ofd-to-pdf "$OUTPUT_DIR/$name.ofd" "$OUTPUT_DIR/$name.pdf"
  dotnet "$CLI" ofd-to-svg "$OUTPUT_DIR/$name.ofd" "$OUTPUT_DIR/$name-1.svg" --pages 1
  dotnet "$CLI" extract-text "$OUTPUT_DIR/$name.ofd" "$OUTPUT_DIR/$name.txt" --include-templates
 done
python3 "$ROOT_DIR/scripts/verify-document-tools-evidence.py" "$OUTPUT_DIR" --render
