#!/usr/bin/env bash
set -euo pipefail
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="${1:?Supply the frozen default feed version}"
DEFAULT_FEED="${2:?Supply the validated default 11-package feed}"
OUTPUT_DIR="${3:-$ROOT_DIR/artifacts/pdf-vectors/package-e2e}"
FONT_DIR="${OFDRW_GRAPHICS_FONTS:-$ROOT_DIR/artifacts/graphics-fonts}"
mkdir -p "$OUTPUT_DIR/feed" "$OUTPUT_DIR/output"
OUTPUT_DIR="$(cd "$OUTPUT_DIR" && pwd)"; DEFAULT_FEED="$(cd "$DEFAULT_FEED" && pwd)"
TASK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/ofd-pdf-vector-consumer.XXXXXX")"
trap 'rm -rf "$TASK_DIR"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$TASK_DIR/dotnet-home}" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
SOURCE_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --nologo)
python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$DEFAULT_FEED" "$VERSION" --verify-manifest
cp "$DEFAULT_FEED"/*.nupkg "$OUTPUT_DIR/feed/"
for product in Ofdrw.Net.Graphics.SkiaSharp Ofdrw.Net.Converter.Pdf.Vector; do
  NUGET_PACKAGES="$SOURCE_CACHE" dotnet pack "$ROOT_DIR/src/$product/$product.csproj" -c Release -o "$OUTPUT_DIR/feed" -p:Version="$VERSION" -p:PackageVersion="$VERSION" "${FLAGS[@]}"
done
cp "$ROOT_DIR/e2e/Ofdrw.Net.Pdf.Vector.E2E/"*.cs "$TASK_DIR/"
cp -R "$ROOT_DIR/e2e/Ofdrw.Net.Pdf.Vector.E2E/testdata" "$TASK_DIR/testdata"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Pdf.Vector.E2E/Ofdrw.Net.Pdf.Vector.E2E.csproj" "$TASK_DIR/Consumer.csproj"
python3 - "$TASK_DIR/NuGet.Config" "$OUTPUT_DIR/feed" "$SOURCE_CACHE" <<'PY'
import sys,xml.etree.ElementTree as ET
root=ET.Element('configuration'); sources=ET.SubElement(root,'packageSources'); ET.SubElement(sources,'clear')
for key,value in [('frozen-artifacts',sys.argv[2]),('cached-third-party',sys.argv[3]),('nuget.org','https://api.nuget.org/v3/index.json')]: ET.SubElement(sources,'add',key=key,value=value)
mapping=ET.SubElement(root,'packageSourceMapping')
for key,pattern in [('frozen-artifacts','Ofdrw.Net.*'),('cached-third-party','*'),('nuget.org','*')]: ET.SubElement(ET.SubElement(mapping,'packageSource',key=key),'package',pattern=pattern)
ET.ElementTree(root).write(sys.argv[1],encoding='utf-8',xml_declaration=True)
PY
export NUGET_PACKAGES="$TASK_DIR/packages"
dotnet restore "$TASK_DIR/Consumer.csproj" --configfile "$TASK_DIR/NuGet.Config" -p:OfdrwPackageVersion="$VERSION" "${FLAGS[@]}"
dotnet build "$TASK_DIR/Consumer.csproj" -c Release --no-restore -p:OfdrwPackageVersion="$VERSION" "${FLAGS[@]}"
dotnet "$TASK_DIR/bin/Release/net10.0/Consumer.dll" "$OUTPUT_DIR/output" "$FONT_DIR/Ofdrw-CI-NotoSansCJKsc-Regular.ttf"
cp "$TASK_DIR/obj/project.assets.json" "$OUTPUT_DIR/consumer.assets.json"
python3 - "$OUTPUT_DIR" "$VERSION" <<'PY'
from pathlib import Path
import hashlib,json,sys,zipfile,xml.etree.ElementTree as ET
root=Path(sys.argv[1]); version=sys.argv[2]; packages=[]
for path in sorted((root/'feed').glob('*.nupkg')):
 with zipfile.ZipFile(path) as z:
  spec=ET.fromstring(z.read(next(n for n in z.namelist() if n.endswith('.nuspec')))); nodes=list(spec.iter())
  find=lambda name:next(n.text for n in nodes if n.tag.rsplit('}',1)[-1]==name)
  if find('version')!=version: raise ValueError('Package version mismatch')
  deps=[n.attrib['id'] for n in nodes if n.tag.rsplit('}',1)[-1]=='dependency']
  if find('id') not in ('Ofdrw.Net.Graphics.SkiaSharp','Ofdrw.Net.Converter.Pdf.Vector') and any('SkiaSharp' in d or d=='Ofdrw.Net.Converter.Pdf.Vector' for d in deps): raise ValueError('Default product acquired optional dependency')
  packages.append({'id':find('id'),'file':path.name,'bytes':path.stat().st_size,'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'dependencies':deps})
if len(packages)!=13: raise ValueError('Expected 11 defaults plus 2 optional packages')
assets=json.loads((root/'consumer.assets.json').read_text())
if any(v['type']=='project' for v in assets['libraries'].values()): raise ValueError('Consumer used project references')
(root/'package-manifest.json').write_text(json.dumps({'version':version,'status':'passed','consumer':'fresh cache, PackageReference only','packages':packages},indent=2)+'\n')
PY
for name in same-source dual vector fallback-source fallback image-absent-source image-absent image-false-source image-false image-true-source image-true image-flags-mixed no-op-source no-op no-content-source no-content closed-point-source closed-point singular-source singular precision-source precision encoding-stream-source encoding-stream cid-map-stream-source cid-map-stream indexed-image-source indexed-image indexed-absent-source indexed-absent indexed-false-source indexed-false indexed-true-source indexed-true indexed-image-clipped-source indexed-image-clipped text-precision-source text-precision; do
  pdftoppm -r 144 -png "$OUTPUT_DIR/output/$name.pdf" "$OUTPUT_DIR/output/$name" >/dev/null 2>&1
done
python3 - "$OUTPUT_DIR/fontconfig.xml" "$FONT_DIR" "$OUTPUT_DIR/font-cache" <<'PY'
import sys,xml.etree.ElementTree as ET
root=ET.Element('fontconfig'); ET.SubElement(root,'dir').text=sys.argv[2]; ET.SubElement(root,'cachedir').text=sys.argv[3]
ET.ElementTree(root).write(sys.argv[1],encoding='utf-8',xml_declaration=True)
PY
for svg in "$OUTPUT_DIR/output/"*.svg; do FONTCONFIG_FILE="$OUTPUT_DIR/fontconfig.xml" rsvg-convert --background-color white -w 840 "$svg" -o "${svg%.svg}-svg.png"; done
python3 "$ROOT_DIR/scripts/verify-pdf-vector-evidence.py" "$OUTPUT_DIR/output"
