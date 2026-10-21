#!/usr/bin/env bash
set -euo pipefail
TASK_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
TASK_OUT="${1:-$TASK_ROOT/artifacts/standard-crypto/package-e2e}"
TASK_VERSION="${2:-0.1.0-r14.local}"
mkdir -p "$TASK_OUT/packages"
TASK_OUT="$(cd "$TASK_OUT" && pwd)"
TASK_TEMP="$(mktemp -d "${TMPDIR:-/tmp}/ofdrw-standard-consume.XXXXXX")"
trap 'rm -rf "$TASK_TEMP"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$TASK_TEMP/dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
TASK_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
TASK_FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false /p:NuGetAudit=false)

# The caller serializes this whole script with run-exclusive.py. Reuse restored assets.
for product in Ofdrw.Net.Core Ofdrw.Net.Packaging Ofdrw.Net.Crypto.Password; do
  dotnet pack "$TASK_ROOT/src/$product/$product.csproj" -c Release --no-restore -o "$TASK_OUT/packages" \
    -p:Version="$TASK_VERSION" -p:PackageVersion="$TASK_VERSION" "${TASK_FLAGS[@]}"
done
python3 - "$TASK_OUT/packages" "$TASK_VERSION" <<'PY'
import hashlib,json,pathlib,sys,xml.etree.ElementTree as ET,zipfile
feed=pathlib.Path(sys.argv[1]);version=sys.argv[2]
expected={f'{n}.{version}.nupkg' for n in ('Ofdrw.Net.Core','Ofdrw.Net.Packaging','Ofdrw.Net.Crypto.Password')}
assert {p.name for p in feed.glob('*.nupkg')}==expected
records=[]
for path in sorted(feed.glob('*.nupkg')):
 with zipfile.ZipFile(path) as z:
  assert z.testzip() is None
  xml=ET.fromstring(z.read(next(x for x in z.namelist() if x.endswith('.nuspec'))))
  values={e.tag.rsplit('}',1)[-1]:e.text for e in xml.iter()}
  dependencies={e.get('id'):e.get('version') for e in xml.iter() if e.tag.rsplit('}',1)[-1]=='dependency'}
  assert values['version']==version
  if values['id']=='Ofdrw.Net.Crypto.Password':
   assert dependencies['BouncyCastle.Cryptography']=='[2.6.2]'
   assert 'lib/net8.0/Ofdrw.Net.Crypto.Password.dll' in z.namelist()
  for dep,dependency_version in dependencies.items():
   if dep.startswith('Ofdrw.Net.'):assert dependency_version in (version,f'[{version}]')
  records.append({'file':path.name,'bytes':path.stat().st_size,'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'dependencies':dependencies})
(feed.parent/'package-manifest.json').write_text(json.dumps({'version':version,'packages':records},indent=2)+'\n')
PY

mkdir -p "$TASK_TEMP/consumer"
cp "$TASK_ROOT/e2e/Ofdrw.Net.Crypto.Standard.E2E/Consumer.cs.txt" "$TASK_TEMP/consumer/Program.cs"
python3 - "$TASK_TEMP" "$TASK_OUT/packages" "$TASK_CACHE" "$TASK_VERSION" <<'PY'
import pathlib,sys,xml.etree.ElementTree as ET
root=pathlib.Path(sys.argv[1]);config=ET.Element('configuration');sources=ET.SubElement(config,'packageSources');ET.SubElement(sources,'clear')
for key,value in [('artifacts',sys.argv[2]),('third-party-cache',sys.argv[3])]:ET.SubElement(sources,'add',key=key,value=value)
mapping=ET.SubElement(config,'packageSourceMapping')
for key,pattern in [('artifacts','Ofdrw.Net.*'),('third-party-cache','*')]:ET.SubElement(ET.SubElement(mapping,'packageSource',key=key),'package',pattern=pattern)
ET.ElementTree(config).write(root/'NuGet.Config',encoding='utf-8',xml_declaration=True)
(root/'consumer/Consumer.csproj').write_text('<Project Sdk="Microsoft.NET.Sdk"><PropertyGroup><TargetFramework>net10.0</TargetFramework><OutputType>Exe</OutputType><ImplicitUsings>enable</ImplicitUsings><Nullable>enable</Nullable></PropertyGroup><ItemGroup><PackageReference Include="Ofdrw.Net.Crypto.Password" Version="'+sys.argv[4]+'" /></ItemGroup></Project>')
PY
export NUGET_PACKAGES="$TASK_TEMP/packages"
dotnet restore "$TASK_TEMP/consumer/Consumer.csproj" --configfile "$TASK_TEMP/NuGet.Config" "${TASK_FLAGS[@]}"
dotnet build "$TASK_TEMP/consumer/Consumer.csproj" -c Release --no-restore "${TASK_FLAGS[@]}"
dotnet "$TASK_TEMP/consumer/bin/Release/net10.0/Consumer.dll" | tee "$TASK_OUT/consumer-result.txt"
cp "$TASK_TEMP/consumer/obj/project.assets.json" "$TASK_OUT/consumer-assets.json"
