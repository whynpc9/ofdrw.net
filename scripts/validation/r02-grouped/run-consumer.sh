#!/usr/bin/env bash
set -euo pipefail
if [[ $# != 4 ]]; then echo "usage: $0 version candidate-feed input-directory fresh-output-directory" >&2; exit 2; fi
root="$(cd "$(dirname "$0")/../../.." && pwd)"
dependency_cache="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
version="$1"; feed="$(cd "$2" && pwd)"; inputs="$(cd "$3" && pwd)"; output="$4"
if [[ -e "$output" ]]; then echo "Output must be a new isolated directory: $output" >&2; exit 2; fi
mkdir -p "$output"; output="$(cd "$output" && pwd)"
NUGET_PACKAGES="$dependency_cache" python3 "$root/scripts/verify-package-artifacts.py" "$feed" "$version" --verify-manifest
cp -R "$root/scripts/validation/r02-grouped/consumer" "$output/consumer"
export DOTNET_CLI_HOME="$output/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
export NUGET_PACKAGES="$output/cache" DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
python3 - "$output/NuGet.Config" "$feed" "$HOME/.nuget/packages" <<'PY'
import sys,xml.etree.ElementTree as E
root=E.Element('configuration');sources=E.SubElement(root,'packageSources');E.SubElement(sources,'clear')
for key,value in [('candidate',sys.argv[2]),('dependencies',sys.argv[3]),('nuget.org','https://api.nuget.org/v3/index.json')]:
 E.SubElement(sources,'add',key=key,value=value)
mapper=E.SubElement(root,'packageSourceMapping')
for key,pattern in [('candidate','Ofdrw.Net.*'),('dependencies','*'),('nuget.org','*')]:
 source=E.SubElement(mapper,'packageSource',key=key);E.SubElement(source,'package',pattern=pattern)
E.ElementTree(root).write(sys.argv[1],encoding='unicode')
PY
flags=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --nologo)
dotnet restore "$output/consumer/Consumer.csproj" --configfile "$output/NuGet.Config" -p:CandidateVersion="$version" -p:NuGetAudit=false "${flags[@]}" > "$output/restore.log" 2>&1
dotnet build "$output/consumer/Consumer.csproj" -c Release --no-restore -p:CandidateVersion="$version" "${flags[@]}" > "$output/build.log" 2>&1
dotnet "$output/consumer/bin/Release/net10.0/Consumer.dll" "$inputs" "$output/result" > "$output/run.log" 2>&1
python3 - "$feed" "$output/result/receipt.json" <<'PY'
import hashlib,json,sys,zipfile
from pathlib import Path
receipt=json.loads(Path(sys.argv[2]).read_text());payloads={}
for package in Path(sys.argv[1]).glob('*.nupkg'):
 with zipfile.ZipFile(package) as z:
  for n in z.namelist():
   if n.startswith('lib/') and n.endswith('.dll'):
    payloads.setdefault(Path(n).stem,set()).add(hashlib.sha256(z.read(n)).hexdigest())
for a in receipt['assemblies']:
 assert a['sha256'] in payloads.get(a['name'],set()),a
receipt['resolvedDllsEqualCandidate']=True
Path(sys.argv[2]).write_text(json.dumps(receipt,indent=2)+'\n')
print('All loaded Ofdrw.Net assemblies match the candidate feed.')
PY
NUGET_PACKAGES="$dependency_cache" python3 "$root/scripts/verify-package-artifacts.py" "$feed" "$version" --verify-manifest
cat "$output/run.log"
