#!/usr/bin/env bash
set -euo pipefail
TASK_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
TASK_OUT="${1:-$TASK_ROOT/artifacts/password-crypto/package-e2e}"
TASK_INPUT="${2:-$TASK_ROOT/artifacts/password-crypto/current}"
TASK_VERSION="${3:-0.1.0-issue11.local}"
if [[ ! "$TASK_VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$ ]]; then
  echo "Expected a SemVer package version." >&2; exit 1
fi
mkdir -p "$TASK_OUT"
TASK_OUT="$(cd "$TASK_OUT" && pwd)"
TASK_TEMP="$(mktemp -d "${TMPDIR:-/tmp}/ofdrw-password-consume.XXXXXX")"
trap 'rm -rf "$TASK_TEMP"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$TASK_TEMP/dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
export DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
TASK_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
export NUGET_PACKAGES="$TASK_CACHE"
TASK_FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false /p:NuGetAudit=false)
mkdir -p "$TASK_OUT/packages"
# Pack this extension and only its own transitive product dependencies.
for product in Ofdrw.Net.Core Ofdrw.Net.Packaging Ofdrw.Net.Crypto.Password; do
  dotnet pack "$TASK_ROOT/src/$product/$product.csproj" -c Release -o "$TASK_OUT/packages" \
    -p:Version="$TASK_VERSION" -p:PackageVersion="$TASK_VERSION" "${TASK_FLAGS[@]}"
done
python3 - "$TASK_OUT/packages" "$TASK_VERSION" "$TASK_ROOT" <<'PY'
import hashlib,json,pathlib,sys,xml.etree.ElementTree as ET,zipfile
feed=pathlib.Path(sys.argv[1]); version=sys.argv[2];repo=pathlib.Path(sys.argv[3])
expected={f'{p}.{version}.nupkg' for p in ['Ofdrw.Net.Core','Ofdrw.Net.Packaging','Ofdrw.Net.Crypto.Password']}
assert {p.name for p in feed.glob('*.nupkg')}==expected
records=[]
for file in sorted(feed.glob('*.nupkg')):
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None
  root=ET.fromstring(z.read(next(p for p in z.namelist() if p.endswith('.nuspec'))))
  nodes={n.tag.rsplit('}',1)[-1]:n.text for n in root.iter()}
  assert nodes['version']==version and nodes['license']=='MIT'
  deps={n.get('id'):n.get('version') for n in root.iter() if n.tag.rsplit('}',1)[-1]=='dependency'}
  if nodes['id']=='Ofdrw.Net.Crypto.Password':
   assert deps['BouncyCastle.Cryptography']=='2.6.2' and 'docs/password-crypto.md' in z.namelist()
   assert 'lib/net8.0/Ofdrw.Net.Crypto.Password.dll' in z.namelist()
  for name,value in deps.items():
   if name.startswith('Ofdrw.Net.'): assert value in [version,f'[{version}]']
  records.append({'file':file.name,'bytes':file.stat().st_size,'sha256':hashlib.sha256(file.read_bytes()).hexdigest(),'dependencies':deps})
# The optional backend must never enter the default product reference graph.
for product in ['Ofdrw.Net.Core','Ofdrw.Net.Packaging','Ofdrw.Net.Converter','Ofdrw.Net.Cli']:
 assert 'BouncyCastle' not in (repo/f'src/{product}/{product}.csproj').read_text()
 assert 'Crypto.Password' not in (repo/f'src/{product}/{product}.csproj').read_text()
(feed/'package-manifest.json').write_text(json.dumps({'version':version,'packages':records,'default_dependencies_unchanged':True},indent=2)+'\n')
PY
mkdir -p "$TASK_TEMP/consumer"
cp "$TASK_ROOT/e2e/Ofdrw.Net.Crypto.Password.E2E/Consumer.cs.txt" "$TASK_TEMP/consumer/Program.cs"
python3 - "$TASK_TEMP" "$TASK_OUT/packages" "$TASK_CACHE" "$TASK_VERSION" <<'PY'
import pathlib,sys,xml.etree.ElementTree as ET
root=pathlib.Path(sys.argv[1]); config=ET.Element('configuration');sources=ET.SubElement(config,'packageSources');ET.SubElement(sources,'clear')
for key,value in [('artifacts',sys.argv[2]),('third-party-cache',sys.argv[3])]:ET.SubElement(sources,'add',key=key,value=value)
mapping=ET.SubElement(config,'packageSourceMapping')
for key,pattern in [('artifacts','Ofdrw.Net.*'),('third-party-cache','*')]:ET.SubElement(ET.SubElement(mapping,'packageSource',key=key),'package',pattern=pattern)
ET.ElementTree(config).write(root/'NuGet.Config',encoding='utf-8',xml_declaration=True)
(root/'consumer/Consumer.csproj').write_text('<Project Sdk="Microsoft.NET.Sdk"><PropertyGroup><TargetFramework>net10.0</TargetFramework><OutputType>Exe</OutputType><ImplicitUsings>enable</ImplicitUsings><Nullable>enable</Nullable></PropertyGroup><ItemGroup><PackageReference Include="Ofdrw.Net.Crypto.Password" Version="'+sys.argv[4]+'" /></ItemGroup></Project>')
PY
export NUGET_PACKAGES="$TASK_TEMP/packages"
dotnet restore "$TASK_TEMP/consumer/Consumer.csproj" --configfile "$TASK_TEMP/NuGet.Config" "${TASK_FLAGS[@]}"
dotnet build "$TASK_TEMP/consumer/Consumer.csproj" -c Release --no-restore "${TASK_FLAGS[@]}"
dotnet "$TASK_TEMP/consumer/bin/Release/net10.0/Consumer.dll" "$TASK_INPUT" "$TASK_OUT/output"
cp "$TASK_TEMP/consumer/obj/project.assets.json" "$TASK_OUT/consumer-assets.json"
echo "[Password E2E] Actual PackageReference consumer and hashes: $TASK_OUT"
