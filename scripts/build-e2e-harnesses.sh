#!/usr/bin/env bash
# Build-only gate: compile the package-referenced e2e harnesses that no other CI step builds,
# against the exact candidate packages. Usage: build-e2e-harnesses.sh VERSION PACKAGES_DIR
# Harnesses are copied under the ignored artifacts/ directory (inheriting the repository's
# Directory.Build.props like the originals), so e2e/ is never modified.
set -euo pipefail
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
if [[ $# -ne 2 ]]; then
  echo "Usage: $0 VERSION PACKAGES_DIR" >&2; exit 2
fi
VERSION="$1"
if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$ ]]; then
  echo "Expected a SemVer package version." >&2; exit 2
fi
PACKAGES_DIR="$(cd "$2" && pwd)"
HARNESSES=(Benchmarks Packaging.Preflight.E2E Converter.Svg.Html.E2E Signatures.Appearance.E2E)

mkdir -p "$ROOT_DIR/artifacts"
TASK_DIR="$(mktemp -d "$ROOT_DIR/artifacts/e2e-harness-build.XXXXXX")"
trap 'rm -rf "$TASK_DIR"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$TASK_DIR/dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
SOURCE_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --nologo)
VERSION_PROPERTIES=(-p:OfdrwPackageVersion="$VERSION" -p:OfdrwAppearancePackageVersion="$VERSION")

python3 - "$TASK_DIR/NuGet.Config" "$PACKAGES_DIR" "$SOURCE_CACHE" <<'PY'
import sys,xml.etree.ElementTree as ET
root=ET.Element('configuration'); sources=ET.SubElement(root,'packageSources'); ET.SubElement(sources,'clear')
mapping=ET.SubElement(root,'packageSourceMapping')
for key,path,pattern in [('candidate',sys.argv[2],'Ofdrw.Net.*'),('cache',sys.argv[3],'*'),('nuget','https://api.nuget.org/v3/index.json','*')]:
    ET.SubElement(sources,'add',key=key,value=path)
    ET.SubElement(ET.SubElement(mapping,'packageSource',key=key),'package',pattern=pattern)
ET.ElementTree(root).write(sys.argv[1],encoding='utf-8',xml_declaration=True)
PY
# A fresh package cache: stale Ofdrw.Net.* packages must not satisfy the restore.
export NUGET_PACKAGES="$TASK_DIR/packages"

built=(); skipped=()
for harness in "${HARNESSES[@]}"; do
  name="Ofdrw.Net.$harness"
  source="$ROOT_DIR/e2e/$name"
  missing=""
  for package in $(grep -o 'PackageReference Include="Ofdrw\.Net[^"]*"' "$source/$name.csproj" | sed 's/.*Include="\([^"]*\)"/\1/'); do
    if [[ ! -f "$PACKAGES_DIR/$package.$VERSION.nupkg" ]]; then missing="$package"; break; fi
  done
  if [[ -n "$missing" ]]; then
    echo "SKIP $name: candidate lacks $missing"; skipped+=("$name"); continue
  fi
  work="$TASK_DIR/projects/$name"
  mkdir -p "$work"
  (cd "$source" && tar --exclude=./bin --exclude=./obj -cf - .) | (cd "$work" && tar -xf -)
  dotnet restore "$work/$name.csproj" --configfile "$TASK_DIR/NuGet.Config" "${VERSION_PROPERTIES[@]}" "${FLAGS[@]}"
  dotnet build "$work/$name.csproj" -c Release --no-restore "${VERSION_PROPERTIES[@]}" "${FLAGS[@]}"
  built+=("$name")
done

echo "Built against $VERSION: ${built[*]:-none}"
echo "Skipped (candidate packages absent): ${skipped[*]:-none}"
