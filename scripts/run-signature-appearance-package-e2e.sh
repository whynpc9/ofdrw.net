#!/usr/bin/env bash
set -euo pipefail
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUTPUT_DIR="${1:-$ROOT_DIR/artifacts/roadmap-v12/R11/package-e2e}"
VERSION="${2:-0.1.0-r11-local.1}"
if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$ ]]; then
  echo "Expected a SemVer package version." >&2; exit 1
fi
mkdir -p "$OUTPUT_DIR"
OUTPUT_DIR="$(cd "$OUTPUT_DIR" && pwd)"
if compgen -G "$OUTPUT_DIR/packages/*.nupkg" > /dev/null; then
  echo "Use a fresh output directory for an immutable package candidate." >&2; exit 1
fi
TASK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/ofd-r11-consumer.XXXXXX")"
trap 'rm -rf "$TASK_DIR"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$OUTPUT_DIR/dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
SOURCE_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --nologo)
mkdir -p "$OUTPUT_DIR/packages" "$TASK_DIR/consumer"
for project in Ofdrw.Net.Core Ofdrw.Net.Packaging Ofdrw.Net.Signatures Ofdrw.Net.Signatures.SesInterop; do
  dotnet pack "$ROOT_DIR/src/$project/$project.csproj" -c Release -o "$OUTPUT_DIR/packages" \
    -p:Version="$VERSION" -p:PackageVersion="$VERSION" "${FLAGS[@]}"
done
python3 - "$OUTPUT_DIR/packages" "$VERSION" <<'PY'
import pathlib,sys,xml.etree.ElementTree as ET,zipfile
folder=pathlib.Path(sys.argv[1]); version=sys.argv[2]
for name in ('Ofdrw.Net.Core','Ofdrw.Net.Packaging','Ofdrw.Net.Signatures','Ofdrw.Net.Signatures.SesInterop'):
    with zipfile.ZipFile(folder/(name+'.'+version+'.nupkg')) as archive:
        root=ET.fromstring(archive.read(name+'.nuspec'))
        for item in root.findall('.//{*}dependency'):
            if item.attrib['id'].startswith('Ofdrw.Net.') and item.attrib['version'].strip('[]') != version:
                raise ValueError(f'{name}: wrong ProjectReference dependency {item.attrib}')
print('Four exact local package identities and dependencies verified.')
PY
cp "$ROOT_DIR/e2e/Ofdrw.Net.Signatures.Appearance.E2E/Program.cs" "$TASK_DIR/consumer/Program.cs"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Signatures.Appearance.E2E/Ofdrw.Net.Signatures.Appearance.E2E.csproj" "$TASK_DIR/consumer/Consumer.csproj"
python3 - "$TASK_DIR/NuGet.Config" "$OUTPUT_DIR/packages" "$SOURCE_CACHE" <<'PY'
import sys,xml.etree.ElementTree as ET
root=ET.Element('configuration'); sources=ET.SubElement(root,'packageSources'); ET.SubElement(sources,'clear')
mapping=ET.SubElement(root,'packageSourceMapping')
for key,path,pattern in [('r11-local',sys.argv[2],'Ofdrw.Net.*'),('cache',sys.argv[3],'*'),('nuget','https://api.nuget.org/v3/index.json','*')]:
    ET.SubElement(sources,'add',key=key,value=path)
    ET.SubElement(ET.SubElement(mapping,'packageSource',key=key),'package',pattern=pattern)
ET.ElementTree(root).write(sys.argv[1],encoding='utf-8',xml_declaration=True)
PY
export NUGET_PACKAGES="$TASK_DIR/packages"
dotnet restore "$TASK_DIR/consumer/Consumer.csproj" --configfile "$TASK_DIR/NuGet.Config" \
  -p:OfdrwAppearancePackageVersion="$VERSION" "${FLAGS[@]}"
dotnet build "$TASK_DIR/consumer/Consumer.csproj" -c Release --no-restore \
  -p:OfdrwAppearancePackageVersion="$VERSION" "${FLAGS[@]}"
dotnet "$TASK_DIR/consumer/bin/Release/net10.0/Consumer.dll" "$OUTPUT_DIR/results"
python3 - "$TASK_DIR/consumer/obj/project.assets.json" "$VERSION" <<'PY'
import json,sys
assets=json.load(open(sys.argv[1])); version=sys.argv[2]
for name in ('Ofdrw.Net.Core','Ofdrw.Net.Packaging','Ofdrw.Net.Signatures','Ofdrw.Net.Signatures.SesInterop'):
    if f'{name}/{version}' not in assets['libraries']: raise ValueError(name+' missing from package consumer')
print('Consumer resolved all four exact local packages.')
PY
