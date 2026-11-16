#!/usr/bin/env bash
# SES dev-interop package consumer.
#
# Usage:
#   run-ses-interop-e2e.sh [OUTPUT_DIR] [VERSION]
#       Standalone: build the solution, pack the default SDK feed and the optional
#       SesInterop package under VERSION, re-run the default package consumer, then
#       run the SES consumer against those local feeds.
#   run-ses-interop-e2e.sh OUTPUT_DIR VERSION --packages-dir DEFAULT_FEED --optional-packages-dir SES_FEED
#       Candidate: consume an existing, manifest-bound default candidate feed and a
#       prepacked Ofdrw.Net.Signatures.SesInterop.VERSION.nupkg. Nothing under src/ is
#       built or packed. The source solution must already be restored (license
#       metadata for verify-package-artifacts.py), as in the Linux package job.
set -euo pipefail
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
POSITIONAL=()
CANDIDATE_FEED=""
OPTIONAL_FEED=""
while [[ $# -gt 0 ]]; do
  case "$1" in
    --packages-dir) CANDIDATE_FEED="${2:?--packages-dir requires a directory}"; shift 2 ;;
    --optional-packages-dir) OPTIONAL_FEED="${2:?--optional-packages-dir requires a directory}"; shift 2 ;;
    -*) echo "Unknown SES interop option: $1" >&2; exit 1 ;;
    *) POSITIONAL+=("$1"); shift ;;
  esac
done
if [[ ${#POSITIONAL[@]} -gt 2 ]]; then
  echo "Usage: $0 [OUTPUT_DIR] [VERSION] [--packages-dir DIR --optional-packages-dir DIR]" >&2; exit 1
fi
OUTPUT_DIR="${POSITIONAL[0]:-$ROOT_DIR/artifacts/ses-interop/e2e}"
VERSION="${POSITIONAL[1]:-0.1.0-ses-interop-test}"
if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$ ]]; then
  echo "Expected a SemVer package version." >&2; exit 1
fi
CANDIDATE_MODE=false
if [[ -n "$CANDIDATE_FEED" || -n "$OPTIONAL_FEED" ]]; then
  if [[ -z "$CANDIDATE_FEED" || -z "$OPTIONAL_FEED" || ${#POSITIONAL[@]} -ne 2 ]]; then
    echo "Candidate mode requires OUTPUT_DIR, VERSION, --packages-dir and --optional-packages-dir." >&2; exit 1
  fi
  for feed in "$CANDIDATE_FEED" "$OPTIONAL_FEED"; do
    if [[ ! -d "$feed" ]]; then echo "Candidate feed directory not found: $feed" >&2; exit 1; fi
  done
  CANDIDATE_FEED="$(cd "$CANDIDATE_FEED" && pwd)"
  OPTIONAL_FEED="$(cd "$OPTIONAL_FEED" && pwd)"
  if [[ ! -f "$CANDIDATE_FEED/package-manifest.json" ]]; then
    echo "Default candidate feed lacks package-manifest.json: $CANDIDATE_FEED" >&2; exit 1
  fi
  CANDIDATE_MODE=true
fi
mkdir -p "$OUTPUT_DIR"
OUTPUT_DIR="$(cd "$OUTPUT_DIR" && pwd)"
if [[ "$CANDIDATE_MODE" == true && -d "$OUTPUT_DIR/packages" ]] && [[ -n "$(find "$OUTPUT_DIR/packages" ! -type d -print -quit)" ]]; then
  # The evidence verifier binds every nupkg under packages/; stale bytes must not join the set.
  echo "Candidate mode requires an evidence package directory without files: $OUTPUT_DIR/packages" >&2; exit 1
fi
TASK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/ofd-ses-consumer.XXXXXX")"
trap 'rm -rf "$TASK_DIR"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$TASK_DIR/dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
SOURCE_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
export NUGET_PACKAGES="$SOURCE_CACHE"
FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --nologo)
mkdir -p "$OUTPUT_DIR/packages/sdk" "$OUTPUT_DIR/packages/optional" "$TASK_DIR/consumer"
if [[ "$CANDIDATE_MODE" == true ]]; then
  # Verify the downloaded candidate against its build-time manifest before use.
  python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$CANDIDATE_FEED" "$VERSION" --verify-manifest
  python3 - "$OPTIONAL_FEED" "$VERSION" "$CANDIDATE_FEED/package-manifest.json" <<'PY_OPTIONAL'
import json, pathlib, sys, xml.etree.ElementTree as E, zipfile
feed, version, manifest = pathlib.Path(sys.argv[1]), sys.argv[2], pathlib.Path(sys.argv[3])
package_id = 'Ofdrw.Net.Signatures.SesInterop'
allowed = {f'{package_id}.{version}.nupkg', f'{package_id}.{version}.snupkg'}
names = {path.name for path in feed.iterdir()}
if f'{package_id}.{version}.nupkg' not in names or not names <= allowed:
    raise SystemExit(f'Optional feed must hold only {package_id}.{version}.nupkg (and its .snupkg); found {sorted(names)} in {feed}')
nupkg = feed / f'{package_id}.{version}.nupkg'
if nupkg.is_symlink():
    raise SystemExit(f'Optional package must be a regular artifact: {nupkg}')
with zipfile.ZipFile(nupkg) as archive:
    if archive.testzip() is not None:
        raise SystemExit(f'Corrupt ZIP payload: {nupkg.name}')
    spec = E.fromstring(archive.read(next(n for n in archive.namelist() if n.endswith('.nuspec'))))
values = {node.tag.rsplit('}', 1)[-1]: node.text for node in spec.iter()}
if values.get('id') != package_id or values.get('version') != version:
    raise SystemExit(f'Optional package identity/version mismatch: {values.get("id")} {values.get("version")}')
default_ids = {item['id'] for item in json.loads(manifest.read_text())['packages']}
siblings = [node for node in spec.iter() if node.tag.rsplit('}', 1)[-1] == 'dependency'
            and node.get('id', '').lower().startswith('ofdrw.net.')]
if not siblings:
    raise SystemExit('Optional package declares no Ofdrw.Net dependency')
for node in siblings:
    if node.get('version') != f'[{version}]' or node.get('id') not in default_ids:
        raise SystemExit(f'Optional package dependency is not an exact candidate package: {node.attrib}')
print(f'Optional {package_id} {version} binds exactly to the validated default candidate.')
PY_OPTIONAL
  # Mirror the exact candidate bytes into the evidence layout the verifier binds.
  find "$CANDIDATE_FEED" -maxdepth 1 -type f \( -name '*.nupkg' -o -name '*.snupkg' -o -name 'package-manifest.json' \) \
    -exec cp -p {} "$OUTPUT_DIR/packages/sdk/" \;
  cp -p "$OPTIONAL_FEED"/* "$OUTPUT_DIR/packages/optional/"
  python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$OUTPUT_DIR/packages/sdk" "$VERSION" --verify-manifest
else
  bash "$ROOT_DIR/scripts/run-converter-package-e2e.sh" "$VERSION" --packages-dir "$OUTPUT_DIR/packages/sdk" --output-dir "$OUTPUT_DIR/sdk-output"
  dotnet pack "$ROOT_DIR/src/Ofdrw.Net.Signatures.SesInterop" -c Release -o "$OUTPUT_DIR/packages/optional" -p:Version="$VERSION" -p:PackageVersion="$VERSION" "${FLAGS[@]}"
fi
cp "$ROOT_DIR/e2e/Ofdrw.Net.Signatures.SesInterop.E2E/Program.cs" "$TASK_DIR/consumer/Program.cs"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Signatures.SesInterop.E2E/Ofdrw.Net.Signatures.SesInterop.E2E.csproj" "$TASK_DIR/consumer/Consumer.csproj"
python3 - "$TASK_DIR/NuGet.Config" "$OUTPUT_DIR/packages/sdk" "$OUTPUT_DIR/packages/optional" "$SOURCE_CACHE" <<'PY'
import sys,xml.etree.ElementTree as E
root=E.Element('configuration'); sources=E.SubElement(root,'packageSources'); E.SubElement(sources,'clear')
mapping=E.SubElement(root,'packageSourceMapping')
for key,path,pattern in [('sdk',sys.argv[2],'Ofdrw.Net.*'),('ses',sys.argv[3],'Ofdrw.Net.Signatures.SesInterop'),('cache',sys.argv[4],'*'),('nuget','https://api.nuget.org/v3/index.json','*')]:
 E.SubElement(sources,'add',key=key,value=path)
 E.SubElement(E.SubElement(mapping,'packageSource',key=key),'package',pattern=pattern)
E.ElementTree(root).write(sys.argv[1],encoding='utf-8',xml_declaration=True)
PY
export NUGET_PACKAGES="$TASK_DIR/packages"
dotnet restore "$TASK_DIR/consumer/Consumer.csproj" --configfile "$TASK_DIR/NuGet.Config" -p:OfdrwSesPackageVersion="$VERSION" "${FLAGS[@]}"
dotnet build "$TASK_DIR/consumer/Consumer.csproj" -c Release --no-restore -p:OfdrwSesPackageVersion="$VERSION" "${FLAGS[@]}"
dotnet tool install Ofdrw.Net.Cli --tool-path "$TASK_DIR/tools" --version "$VERSION" --configfile "$TASK_DIR/NuGet.Config"
FONT_DIR="${OFDRW_SES_FONTS:-$ROOT_DIR/artifacts/ses-interop/fonts}"
dotnet "$TASK_DIR/consumer/bin/Release/net10.0/Consumer.dll" "$ROOT_DIR" "$OUTPUT_DIR/files" "$FONT_DIR"
for mode in native default; do
  for version in v1 v4; do
    name="$mode-$version"
    set +e
    "$TASK_DIR/tools/ofdrw" verify-signatures "$OUTPUT_DIR/files/$name.ofd" > "$OUTPUT_DIR/files/$name-cli.txt" 2>&1
    result=$?
    set -e
    if [[ "$result" -ne 2 ]]; then cat "$OUTPUT_DIR/files/$name-cli.txt"; echo "Expected CLI reference-only exit 2, got $result" >&2; exit 1; fi
  done
done
python3 "$ROOT_DIR/scripts/verify-ses-interop-evidence.py" "$OUTPUT_DIR" "$VERSION" "$TASK_DIR/consumer/obj/project.assets.json" "$TASK_DIR/tools/.store/ofdrw.net.cli/$VERSION/ofdrw.net.cli/$VERSION/tools/net10.0/any/Ofdrw.Net.Cli.deps.json"
NUGET_PACKAGES="$SOURCE_CACHE" python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$OUTPUT_DIR/packages/sdk" "$VERSION" --verify-manifest
if [[ "$CANDIDATE_MODE" == true ]]; then
  # The candidate must still match its manifest after use.
  NUGET_PACKAGES="$SOURCE_CACHE" python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$CANDIDATE_FEED" "$VERSION" --verify-manifest
  python3 - "$OUTPUT_DIR" "$VERSION" "$ROOT_DIR" <<'PY_SOURCE'
import hashlib, json, pathlib, subprocess, sys
root, version, repo = pathlib.Path(sys.argv[1]), sys.argv[2], sys.argv[3]
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
(root / 'candidate-source.json').write_text(json.dumps({
    'mode': 'candidate', 'version': version,
    'sourceSha': subprocess.check_output(['git', '-C', repo, 'rev-parse', 'HEAD'], text=True).strip(),
    'defaultManifestSha256': sha(root / 'packages/sdk/package-manifest.json'),
    'optionalPackageSha256': sha(root / f'packages/optional/Ofdrw.Net.Signatures.SesInterop.{version}.nupkg'),
    'defaultConsumerEvidence': 'Not re-run here; the Linux package-e2e job consumes the same manifest-bound candidate.',
}, indent=2) + '\n')
PY_SOURCE
fi
