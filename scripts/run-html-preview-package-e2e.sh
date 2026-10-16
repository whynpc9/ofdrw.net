#!/usr/bin/env bash
set -euo pipefail
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="${1:?local package version required}"
FEED="${2:?local package directory required}"
RESULT="${3:?output directory required}"
mkdir -p "$RESULT"
FEED="$(cd "$FEED" && pwd)"
RESULT="$(cd "$RESULT" && pwd)"
TASK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/ofdrw-html-consumer.XXXXXX")"
trap 'rm -rf "$TASK_DIR"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$TASK_DIR/dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1
SOURCE_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
python3 - "$TASK_DIR/NuGet.Config" "$FEED" "$SOURCE_CACHE" <<'PY'
import sys,xml.etree.ElementTree as E
root=E.Element('configuration');sources=E.SubElement(root,'packageSources');E.SubElement(sources,'clear')
for key,value in [('local-artifacts',sys.argv[2]),('cached-third-party',sys.argv[3])]:E.SubElement(sources,'add',key=key,value=value)
mapping=E.SubElement(root,'packageSourceMapping')
for key,pattern in [('local-artifacts','Ofdrw.Net.*'),('cached-third-party','*')]:E.SubElement(E.SubElement(mapping,'packageSource',key=key),'package',pattern=pattern)
E.ElementTree(root).write(sys.argv[1],encoding='utf-8',xml_declaration=True)
PY
mkdir -p "$TASK_DIR/consumer"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Svg.Html.E2E/Program.cs" "$TASK_DIR/consumer/Program.cs"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Svg.Html.E2E/Ofdrw.Net.Converter.Svg.Html.E2E.csproj" "$TASK_DIR/consumer/Consumer.csproj"
export NUGET_PACKAGES="$TASK_DIR/packages"
FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --nologo)
dotnet restore "$TASK_DIR/consumer/Consumer.csproj" --configfile "$TASK_DIR/NuGet.Config" -p:OfdrwPackageVersion="$VERSION" "${FLAGS[@]}"
dotnet build "$TASK_DIR/consumer/Consumer.csproj" -c Release --no-restore -p:OfdrwPackageVersion="$VERSION" "${FLAGS[@]}"
dotnet run --project "$TASK_DIR/consumer/Consumer.csproj" -c Release --no-build --no-restore -- "$RESULT"
dotnet tool install Ofdrw.Net.Cli --tool-path "$TASK_DIR/tools" --version "$VERSION" --configfile "$TASK_DIR/NuGet.Config"
"$TASK_DIR/tools/ofdrw" ofd-to-html "$RESULT/package-three-pages.ofd" "$RESULT/package-cli.html" --pages 3,1,3 --title 'Installed CLI <literal> & title'
python3 - "$RESULT/package-cli.html" <<'PY'
import sys,xml.etree.ElementTree as E
r=E.parse(sys.argv[1]).getroot()
assert [s.attrib['data-source-page'] for s in r.iter('section')]==['3','1','3']
assert r.find('head/title').text=='Installed CLI <literal> & title'
print('Installed local-package CLI HTML order and escaping passed.')
PY
