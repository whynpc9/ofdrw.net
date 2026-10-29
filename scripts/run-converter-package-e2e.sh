#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="${1:-0.1.0-preview.8}"
if [[ $# -gt 0 ]]; then shift; fi
if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$ ]]; then
  echo "Expected a SemVer package version." >&2
  exit 1
fi
OUT_DIR="$ROOT_DIR/artifacts/package-e2e/$VERSION/packages"
RESULT_DIR="$ROOT_DIR/artifacts/package-e2e/$VERSION/output"
CONSUME_ONLY=false
PACK_ONLY=false
NO_BUILD=false
while [[ $# -gt 0 ]]; do
  case "$1" in
    --consume-only) CONSUME_ONLY=true; shift ;;
    --pack-only) PACK_ONLY=true; shift ;;
    --no-build) NO_BUILD=true; shift ;;
    --packages-dir) OUT_DIR="$2"; shift 2 ;;
    --output-dir) RESULT_DIR="$2"; shift 2 ;;
    *) echo "Unknown package E2E option: $1" >&2; exit 1 ;;
  esac
done
if [[ "$PACK_ONLY" == true && "$CONSUME_ONLY" == true ]] || [[ "$NO_BUILD" == true && "$PACK_ONLY" != true ]]; then
  echo "--pack-only and --consume-only are exclusive; --no-build requires --pack-only." >&2
  exit 1
fi
if [[ "$PACK_ONLY" == true && -d "$OUT_DIR" ]] && [[ -n "$(find "$OUT_DIR" -mindepth 1 -maxdepth 1 -print -quit)" ]]; then
  echo "Pack-only requires an empty candidate directory: $OUT_DIR" >&2
  exit 1
fi
mkdir -p "$OUT_DIR" "$RESULT_DIR"
OUT_DIR="$(cd "$OUT_DIR" && pwd)"
RESULT_DIR="$(cd "$RESULT_DIR" && pwd)"
TASK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/ofdrw-package-e2e.XXXXXX")"
trap 'rm -rf "$TASK_DIR"' EXIT
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$TASK_DIR/dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1
export DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1
SOURCE_CACHE="${NUGET_PACKAGES:-$HOME/.nuget/packages}"
# DOTNET_CLI_HOME also changes NuGet's default cache on clean runners. Keep
# the build cache and the consumer's third-party feed on the same explicit path.
mkdir -p "$SOURCE_CACHE"
SOURCE_CACHE="$(cd "$SOURCE_CACHE" && pwd)"
export NUGET_PACKAGES="$SOURCE_CACHE"
BUILD_FLAGS=(--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --nologo)

if [[ "$CONSUME_ONLY" == false ]]; then
  if [[ "$NO_BUILD" == false ]]; then
    dotnet restore "$ROOT_DIR/Ofdrw.Net.sln" "${BUILD_FLAGS[@]}"
    dotnet build "$ROOT_DIR/Ofdrw.Net.sln" -c Release --no-restore \
      -p:Version="$VERSION" -p:PackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
  fi
  # The default converter/tool feed follows the validator's current product
  # set, including the optional HarfBuzz text adapter. Other optional
  # extensions have independent consumption gates.
  python3 - "$ROOT_DIR/scripts" > "$TASK_DIR/default-products.txt" <<'PY_PACKAGES'
import sys
sys.path.insert(0, sys.argv[1])
from importlib import import_module
print('\n'.join(import_module('verify-package-artifacts').PACKAGES))
PY_PACKAGES
  while IFS= read -r product; do
    dotnet pack "$ROOT_DIR/src/$product/$product.csproj" -c Release --no-build --no-restore -o "$OUT_DIR" \
      -p:Version="$VERSION" -p:PackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
  done < "$TASK_DIR/default-products.txt"
fi
VALIDATION=(python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$OUT_DIR" "$VERSION")
if [[ "$CONSUME_ONLY" == true && -f "$OUT_DIR/package-manifest.json" ]]; then
  VALIDATION+=(--verify-manifest)
fi
"${VALIDATION[@]}"
if [[ "$PACK_ONLY" == true ]]; then
  echo "[E2E] Produced validated candidate packages: $OUT_DIR"
  exit 0
fi

# A clean consumer directory and package cache prevent old same-version DLLs or
# project references from satisfying the test. Only this artifact feed may supply
# Ofdrw.Net packages; third-party dependencies may reuse the existing local feed.
mkdir -p "$TASK_DIR/consumer"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs" "$TASK_DIR/consumer/Program.cs"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Pdf.E2E/Ofdrw.Net.Converter.Pdf.E2E.csproj" "$TASK_DIR/consumer/Consumer.csproj"
python3 - "$TASK_DIR/NuGet.Config" "$OUT_DIR" "$SOURCE_CACHE" <<'PY'
import sys
import xml.etree.ElementTree as ET
root = ET.Element('configuration')
sources = ET.SubElement(root, 'packageSources')
ET.SubElement(sources, 'clear')
for key, value in [('local-artifacts', sys.argv[2]), ('cached-third-party', sys.argv[3]), ('nuget.org', 'https://api.nuget.org/v3/index.json')]:
    ET.SubElement(sources, 'add', key=key, value=value)
mapping = ET.SubElement(root, 'packageSourceMapping')
for key, pattern in [('local-artifacts', 'Ofdrw.Net.*'), ('cached-third-party', '*'), ('nuget.org', '*')]:
    source = ET.SubElement(mapping, 'packageSource', key=key)
    ET.SubElement(source, 'package', pattern=pattern)
ET.ElementTree(root).write(sys.argv[1], encoding='utf-8', xml_declaration=True)
PY
export NUGET_PACKAGES="$TASK_DIR/packages"
CONSUMER="$TASK_DIR/consumer/Consumer.csproj"
dotnet restore "$CONSUMER" --configfile "$TASK_DIR/NuGet.Config" -p:OfdrwPackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
dotnet build "$CONSUMER" -c Release --no-restore -p:OfdrwPackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
dotnet tool install Ofdrw.Net.Cli --tool-path "$TASK_DIR/tools" --version "$VERSION" --configfile "$TASK_DIR/NuGet.Config"
python3 - "$RESULT_DIR/cli-plain-text.txt" <<'PY_TEXT'
import pathlib, sys
pathlib.Path(sys.argv[1]).write_bytes("\ufeff中英 iii WWW\r\n\r\n尾行\n".encode('utf-8'))
PY_TEXT
"$TASK_DIR/tools/ofdrw" text-to-ofd "$RESULT_DIR/cli-plain-text.txt" "$RESULT_DIR/cli-plain-text.ofd"
"$TASK_DIR/tools/ofdrw" extract-text "$RESULT_DIR/cli-plain-text.ofd" "$RESULT_DIR/cli-plain-text-extracted.txt"
python3 - "$RESULT_DIR/cli-plain-text-extracted.txt" <<'PY_TEXT_CHECK'
import pathlib, sys
text = pathlib.Path(sys.argv[1]).read_text()
assert text.count('中英 iii WWW') == 1 and text.count('尾行') == 1, repr(text)
PY_TEXT_CHECK
"$TASK_DIR/tools/ofdrw" docx-to-ofd \
  "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx" "$RESULT_DIR/cli-native.ofd"
"$TASK_DIR/tools/ofdrw" ofd-to-image "$RESULT_DIR/cli-native.ofd" "$RESULT_DIR/cli-image.png" --pages 1 --ppm 2
"$TASK_DIR/tools/ofdrw" image-to-ofd "$RESULT_DIR/cli-image.png" "$RESULT_DIR/cli-image.ofd" --ppm 2 --page-width 220 --page-height 307
OFDRW_REPO_ROOT="$ROOT_DIR" OFDRW_E2E_OUTPUT_DIR="$RESULT_DIR" dotnet run \
  --project "$CONSUMER" -c Release --no-build --no-restore
# Exercise the production subset backend from the newly packed SDK as well.
mkdir -p "$TASK_DIR/font-consumer"
cp "$ROOT_DIR/e2e/Ofdrw.Net.FontSubset.E2E/Program.cs" "$TASK_DIR/font-consumer/Program.cs"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Converter.Pdf.E2E/Ofdrw.Net.Converter.Pdf.E2E.csproj" "$TASK_DIR/font-consumer/Consumer.csproj"
FONT_CONSUMER="$TASK_DIR/font-consumer/Consumer.csproj"
dotnet restore "$FONT_CONSUMER" --configfile "$TASK_DIR/NuGet.Config" -p:OfdrwPackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
dotnet build "$FONT_CONSUMER" -c Release --no-restore -p:OfdrwPackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
dotnet run --project "$FONT_CONSUMER" -c Release --no-build --no-restore -- "$ROOT_DIR" "$RESULT_DIR/font-subset"

# The release feed includes the optional HarfBuzz adapter. Restore and run the
# fixed-font page probe from packages only, so missing native assets or a bad
# adapter dependency cannot pass the candidate gate on project references.
mkdir -p "$TASK_DIR/text-consumer"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Text.HarfBuzz.E2E/Program.cs" "$TASK_DIR/text-consumer/Program.cs"
cat > "$TASK_DIR/text-consumer/Consumer.csproj" <<EOF_TEXT_CONSUMER
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup><OutputType>Exe</OutputType><TargetFramework>net10.0</TargetFramework><ImplicitUsings>enable</ImplicitUsings></PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Ofdrw.Net.Text.HarfBuzz" Version="$VERSION" />
    <PackageReference Include="Ofdrw.Net.Packaging" Version="$VERSION" />
    <PackageReference Include="Ofdrw.Net.Reader" Version="$VERSION" />
    <PackageReference Include="Ofdrw.Net.Converter.Pdf" Version="$VERSION" />
    <PackageReference Include="Ofdrw.Net.Converter.Svg" Version="$VERSION" />
  </ItemGroup>
</Project>
EOF_TEXT_CONSUMER
TEXT_CONSUMER="$TASK_DIR/text-consumer/Consumer.csproj"
TEXT_OUTPUT="$RESULT_DIR/text-prelayout"
dotnet restore "$TEXT_CONSUMER" --configfile "$TASK_DIR/NuGet.Config" "${BUILD_FLAGS[@]}"
dotnet build "$TEXT_CONSUMER" -c Release --no-restore "${BUILD_FLAGS[@]}"
mkdir -p "$TEXT_OUTPUT"
python3 - "$TEXT_CONSUMER" "$TASK_DIR/packages" "$OUT_DIR" "$VERSION" "$TEXT_OUTPUT/text-consumer-identities.json" "$ROOT_DIR" <<'PY_TEXT_IDENTITY'
import hashlib, json, pathlib, sys, zipfile, subprocess
project, cache, feed, version, receipt, root = map(pathlib.Path, sys.argv[1:])
version = str(version)
bin_dir = project.parent / 'bin/Release/net10.0'
deps = json.loads((bin_dir / 'Consumer.deps.json').read_text())
target = deps['targets'][deps['runtimeTarget']['name']]
manifest = {item['id']: item for item in json.loads((feed / 'package-manifest.json').read_text())['packages']}
sha = lambda data: hashlib.sha256(data).hexdigest()
items = []
for key, node in sorted(target.items()):
    if not key.startswith('Ofdrw.Net.'):
        continue
    package_id, resolved_version = key.split('/', 1)
    assert resolved_version == version and package_id in manifest, key
    assets = [path for path in node.get('runtime', {}) if path.endswith('.dll')]
    assert len(assets) == 1, (key, assets)
    asset = assets[0]
    nupkg = feed / manifest[package_id]['file']
    assert sha(nupkg.read_bytes()) == manifest[package_id]['sha256'], key
    with zipfile.ZipFile(nupkg) as archive:
        payload = archive.read(asset)
    extracted = cache / package_id.lower() / version / asset
    copied = bin_dir / pathlib.Path(asset).name
    assert sha(payload) == sha(extracted.read_bytes()) == sha(copied.read_bytes()), key
    items.append({'id': package_id, 'version': version, 'asset': asset,
        'nupkgSha256': manifest[package_id]['sha256'], 'dllSha256': sha(payload)})
assert {'Ofdrw.Net.Core', 'Ofdrw.Net.Layout', 'Ofdrw.Net.Text.HarfBuzz',
    'Ofdrw.Net.Packaging', 'Ofdrw.Net.Reader', 'Ofdrw.Net.Converter.Pdf',
    'Ofdrw.Net.Converter.Svg'} <= {item['id'] for item in items}
receipt.write_text(json.dumps({'sourceSha': subprocess.check_output(
    ['git', '-C', str(root), 'rev-parse', 'HEAD'], text=True).strip(),
    'runtimeTarget': deps['runtimeTarget']['name'], 'packages': items}, indent=2) + '\n')
PY_TEXT_IDENTITY
(cd "$ROOT_DIR" && R21_SOURCE_SHA="$(git rev-parse HEAD)" dotnet run \
  --project "$TEXT_CONSUMER" -c Release --no-build --no-restore -- "$TEXT_OUTPUT")
for sample in ascii han; do
  pdftoppm -f 1 -l 1 -png -singlefile -r 144 \
    "$TEXT_OUTPUT/r21-$sample.pdf" "$TEXT_OUTPUT/r21-$sample-pdf-page"
  test -s "$TEXT_OUTPUT/r21-$sample.ofd"
  test -s "$TEXT_OUTPUT/r21-$sample.png"
  test -s "$TEXT_OUTPUT/r21-$sample-pdf-page.png"
done
python3 - "$TEXT_OUTPUT/r21-receipt.json" <<'PY_TEXT_CONSUMER'
import json, sys
receipt = json.load(open(sys.argv[1]))
assert receipt['ofdText'] == receipt['ascii']
assert receipt['han']['ofdText'] == '中\n文\nA\n中'
assert 'TEXT_FACE_FALLBACK' in receipt['han']['fallback']
assert all(output['bytes'] > 0 for output in receipt['outputs'])
PY_TEXT_CONSUMER
python3 - "$TEXT_OUTPUT/r22-receipt.json" <<'PY_R22_TEXT_CONSUMER'
import json, sys
receipt = json.load(open(sys.argv[1]))
assert {item['source'] for item in receipt['cases']} == {'AV', 'ffi', 'e\u0301', 'a\u0308'}
assert {item['feature'] for item in receipt['cases'] if item['source'] == 'AV'} == {'on', 'off'}
assert {item['feature'] for item in receipt['cases'] if item['source'] == 'ffi'} == {'on', 'off'}
av = {item['feature']: item for item in receipt['cases'] if item['source'] == 'AV'}
assert av['on']['origins'][1]['OriginXMillimeters'] != av['off']['origins'][1]['OriginXMillimeters']
acute = next(item for item in receipt['cases'] if item['name'] == 'e-acute')
assert acute['source'] == 'e\u0301' and len(acute['shapedGids']) == 1
diaeresis = next(item for item in receipt['cases'] if item['name'] == 'a-diaeresis')
assert diaeresis['source'] == 'a\u0308' and len(diaeresis['shapedGids']) == 1
assert all(item['readback'] == item['source'] for item in receipt['cases'])
assert all(item['svgGids'] == item['shapedGids'] for item in receipt['cases'])
assert all(output['bytes'] > 0 for output in receipt['outputs'])
PY_R22_TEXT_CONSUMER
python3 - "$TEXT_OUTPUT/r22-mixed-query-receipt.json" "$TEXT_OUTPUT/r22-mixed-output-receipt.json" <<'PY_R22_MIXED_CONSUMER'
import json, sys
query, output = (json.load(open(path)) for path in sys.argv[1:])
source = 'abc \u05e9\u05dc\u05d5\u05dd def'
assert query['source'] == output['source'] == source
assert query['sourceSha'] == output['sourceSha']
assert isinstance(query['sourceSha'], str) and len(query['sourceSha']) == 40
assert query['face']['FaceSha256'] == output['face']
wide = query['wide'][0]['runs']
assert [run['Text'] for run in wide] == ['abc', ' ', '\u05e9\u05dc\u05d5\u05dd', ' ', 'def']
assert [glyph['ClusterStart'] for glyph in wide[2]['glyphs']] == [7, 6, 5, 4]
assert [line['source'] for line in query['wrapped']] == ['abc ', '\u05e9\u05dc\u05d5\u05dd ', 'def']
two = query['twoWord']
assert two['source'] == 'abc \u05e9\u05dc\u05d5\u05dd \u05de\u05e7\u05d5\u05dd def'
assert [run['Text'] for run in two['runs']] == ['abc', ' ', '\u05e9\u05dc\u05d5\u05dd', ' ',
    '\u05de\u05e7\u05d5\u05dd', ' ', 'def']
assert two['runs'][4]['XMillimeters'] < two['runs'][2]['XMillimeters']
assert {item['name'] for item in output['outputs']} == {
    f'r22-mixed-{kind}.{extension}'
    for kind in ('wide', 'short-words', 'long-upper', 'unequal-latin', 'upper-leading', 'exact-fit')
    for extension in ('ofd', 'pdf', 'svg', 'png')}
assert {item['name'] for item in output['cases']} == {
    f'r22-mixed-{kind}' for kind in
    ('wide', 'short-words', 'long-upper', 'unequal-latin', 'upper-leading', 'exact-fit')}
assert all(item['bytes'] > 0 for item in output['outputs'])
PY_R22_MIXED_CONSUMER
for sample in av av-kern-off ffi-on ffi-off e-acute a-diaeresis; do
  pdftoppm -f 1 -l 1 -png -singlefile -r 144 \
    "$TEXT_OUTPUT/r22-$sample.pdf" "$TEXT_OUTPUT/r22-$sample-pdf-page"
  test -s "$TEXT_OUTPUT/r22-$sample-pdf-page.png"
done
for sample in wide short-words long-upper unequal-latin upper-leading exact-fit; do
  pdftoppm -f 1 -l 1 -png -singlefile -r 144 \
    "$TEXT_OUTPUT/r22-mixed-$sample.pdf" "$TEXT_OUTPUT/r22-mixed-$sample-pdf-page"
  test -s "$TEXT_OUTPUT/r22-mixed-$sample-pdf-page.png"
done
python3 - "$TEXT_OUTPUT" <<'PY_R22_PDF_TEXT'
import json, pathlib, subprocess, sys
root = pathlib.Path(sys.argv[1])
for case in json.loads((root / 'r22-receipt.json').read_text())['cases']:
    pdf = str(root / ('r22-' + case['name'] + '.pdf'))
    extracted = subprocess.check_output(['pdftotext', '-enc', 'UTF-8', pdf, '-']).decode().strip()
    assert extracted == case['source'], (case['name'], extracted)
    fonts = subprocess.check_output(['pdffonts', pdf], text=True).splitlines()[2:]
    assert len(fonts) == 1 and 'CID TrueType' in fonts[0] and 'yes yes yes' in fonts[0], case['name']
PY_R22_PDF_TEXT
# Exercise public complex-table pagination from the exact candidate packages.
# The source E2E producer is copied into a fresh PackageReference-only project;
# package-source mapping and the private cache above prevent project-reference reuse.
mkdir -p "$TASK_DIR/table-consumer"
cp "$ROOT_DIR/e2e/Ofdrw.Net.Layout.E2E/ComplexTableSamples.cs" "$TASK_DIR/table-consumer/ComplexTableSamples.cs"
cat > "$TASK_DIR/table-consumer/Program.cs" <<'EOF_TABLE_PROGRAM'
await ComplexTableSamples.Run(args[0]);
EOF_TABLE_PROGRAM
cat > "$TASK_DIR/table-consumer/Consumer.csproj" <<EOF_TABLE_CONSUMER
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup><OutputType>Exe</OutputType><TargetFramework>net10.0</TargetFramework><ImplicitUsings>enable</ImplicitUsings></PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Ofdrw.Net.Layout" Version="$VERSION" />
    <PackageReference Include="Ofdrw.Net.Packaging" Version="$VERSION" />
    <PackageReference Include="Ofdrw.Net.Reader" Version="$VERSION" />
    <PackageReference Include="Ofdrw.Net.Converter.Pdf" Version="$VERSION" />
  </ItemGroup>
</Project>
EOF_TABLE_CONSUMER
TABLE_CONSUMER="$TASK_DIR/table-consumer/Consumer.csproj"
TABLE_OUTPUT="$RESULT_DIR/complex-table"
mkdir -p "$TABLE_OUTPUT"
dotnet restore "$TABLE_CONSUMER" --configfile "$TASK_DIR/NuGet.Config" "${BUILD_FLAGS[@]}"
dotnet build "$TABLE_CONSUMER" -c Release --no-restore "${BUILD_FLAGS[@]}"
R19_SOURCE_SHA="$(git -C "$ROOT_DIR" rev-parse HEAD)" dotnet run \
  --project "$TABLE_CONSUMER" -c Release --no-build --no-restore -- "$TABLE_OUTPUT"
# Verify the public attachment-removal API through the same exact local feed.
mkdir -p "$TASK_DIR/attachment-consumer"
cp "$ROOT_DIR/e2e/Ofdrw.Net.AttachmentCleanup.E2E/Program.cs" "$TASK_DIR/attachment-consumer/Program.cs"
cp "$ROOT_DIR/e2e/Ofdrw.Net.AttachmentCleanup.E2E/Ofdrw.Net.AttachmentCleanup.E2E.csproj" "$TASK_DIR/attachment-consumer/Consumer.csproj"
ATTACHMENT_CONSUMER="$TASK_DIR/attachment-consumer/Consumer.csproj"
dotnet restore "$ATTACHMENT_CONSUMER" --configfile "$TASK_DIR/NuGet.Config" \
  -p:R18PackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
dotnet build "$ATTACHMENT_CONSUMER" -c Release --no-restore \
  -p:R18PackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
R18_SOURCE_SHA="$(git -C "$ROOT_DIR" rev-parse HEAD)" R18_E2E_OUTPUT_DIR="$RESULT_DIR/attachment-cleanup" \
  dotnet "$TASK_DIR/attachment-consumer/bin/Release/net10.0/Consumer.dll" | tee "$RESULT_DIR/attachment-cleanup-consumer.txt"
NUGET_PACKAGES="$SOURCE_CACHE" python3 "$ROOT_DIR/scripts/verify-package-artifacts.py" "$OUT_DIR" "$VERSION" --verify-manifest
cp "$OUT_DIR/package-manifest.json" "$RESULT_DIR/package-manifest.json"
echo "[E2E] Verified package bytes and output: $RESULT_DIR"
