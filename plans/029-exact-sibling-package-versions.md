# Plan 029: Ofdrw.Net packages depend on the exact same version of each other

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- Directory.Build.props Directory.Build.targets scripts/verify-package-artifacts.py scripts/tests/test_package_artifacts.py src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj`
> If any of these changed since this plan was written (or
> `Directory.Build.targets` now exists), compare the "Current state" excerpts
> against the live files before proceeding; on a mismatch, treat it as a STOP
> condition.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: MED (changes every published package's dependency metadata)
- **Depends on**: none
- **Category**: bug, dependencies
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

The SDK ships as twelve NuGet packages (`Ofdrw.Net.Core`, `.Packaging`,
`.Layout`, `.Reader`, …). They call each other's **internal** members: the
shipping projects grant each other about twenty `InternalsVisibleTo` accesses
(for example `Ofdrw.Net.Core` to `Packaging`, `Signatures`, `Reader`, `Layout`,
`Converter.Docx`, `Converter.Pdf`, `Converter.Svg`). Internal members carry no
compatibility promise between releases.

Yet each package declares its siblings as *minimum* versions. Packed at
`427433a`, `Ofdrw.Net.Reader.0.3.1.nupkg` contains:

```xml
<dependency id="Ofdrw.Net.Core" version="0.3.1" exclude="Build,Analyzers" />
<dependency id="Ofdrw.Net.Packaging" version="0.3.1" exclude="Build,Analyzers" />
```

`version="0.3.1"` means "0.3.1 or newer". An application that references
`Ofdrw.Net.Reader 0.3.1` and later adds `Ofdrw.Net.Core 0.4.0` restores without
any warning and can fail at run time with `MissingMethodException` or
`TypeLoadException` when Reader calls a Core internal that moved.

After this plan every `Ofdrw.Net.*` → `Ofdrw.Net.*` dependency is the exact
range `[x.y.z]`, NuGet reports the mismatch at restore time, and the package
validator refuses a package set that is not pinned.

## Current state

### How packages are versioned and packed

`Directory.Build.props` (repo root) sets the version for every project:

```xml
    <VersionPrefix Condition="'$(VersionPrefix)' == ''">0.3.1</VersionPrefix>
    <PackageVersion Condition="'$(PackageVersion)' == '' and '$(VersionSuffix)' != ''">$(VersionPrefix)-$(VersionSuffix)</PackageVersion>
    <PackageVersion Condition="'$(PackageVersion)' == ''">$(VersionPrefix)</PackageVersion>
    <Version Condition="'$(Version)' == ''">$(PackageVersion)</Version>
```

There is no `Directory.Build.targets` in the repository. Siblings are
`ProjectReference`s, e.g. `src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj`:

```xml
  <ItemGroup>
    <InternalsVisibleTo Include="Ofdrw.Net.Converter.Pdf" />
    <InternalsVisibleTo Include="Ofdrw.Net.Packaging.Tests" />
    <InternalsVisibleTo Include="Ofdrw.Net.Cli" />
    <ProjectReference Include="../Ofdrw.Net.Reader/Ofdrw.Net.Reader.csproj" />
    <ProjectReference Include="../Ofdrw.Net.Packaging/Ofdrw.Net.Packaging.csproj" />
    <ProjectReference Include="../Ofdrw.Net.Core/Ofdrw.Net.Core.csproj" />
  </ItemGroup>
  <ItemGroup>
    <InternalsVisibleTo Include="Ofdrw.Net.Converter.Docx" />
    <InternalsVisibleTo Include="Ofdrw.Net.Converter.Pdf" />
    <InternalsVisibleTo Include="Ofdrw.Net.Layout.Tests" />
  </ItemGroup>
```

(`Ofdrw.Net.Converter.Pdf` is granted twice — lines 8 and 17. Harmless, removed
in Step 4.)

Every pack in the repository passes the version explicitly, e.g.
`.github/workflows/publish-nuget.yml:144` and
`scripts/run-converter-package-e2e.sh:67`:

```sh
    dotnet pack "$ROOT_DIR/src/$product/$product.csproj" -c Release --no-build --no-restore -o "$OUT_DIR" \
      -p:Version="$VERSION" -p:PackageVersion="$VERSION" "${BUILD_FLAGS[@]}"
```

### The validator accepts both forms

`scripts/verify-package-artifacts.py:51-54`:

```python
            for dependency in root.iter():
                if dependency.tag.rsplit('}', 1)[-1] == 'dependency' and dependency.get('id', '').startswith('Ofdrw.Net.'):
                    if dependency.get('version') not in (version, f'[{version}]'):
                        raise ValueError(f'Mismatched SDK dependency in {path.name}: {dependency.attrib}')
```

Its unit tests are in `scripts/tests/test_package_artifacts.py`. The fixture
packages built in `setUp` have no `<dependencies>` at all; the one dependency
test rewrites a nuspec (lines 60–67):

```python
    def test_wrong_sdk_dependency_is_rejected(self):
        path = self.root / f'Ofdrw.Net.Core.{self.version}.nupkg'
        with zipfile.ZipFile(path) as archive:
            contents = {name: archive.read(name) for name in archive.namelist()}
        contents['Ofdrw.Net.Core.nuspec'] = f'<package><metadata><id>Ofdrw.Net.Core</id><version>{self.version}</version><license type="expression">MIT</license><dependencies><dependency id="Ofdrw.Net.Reader" version="0.0.0" /></dependencies></metadata></package>'.encode()
        with zipfile.ZipFile(path, 'w') as archive:
            for name, data in contents.items(): archive.writestr(name, data)
        with self.assertRaises(ValueError): packages.inspect(self.root, self.version)
```

### The mechanism was tried

During planning this target was injected from outside the repository (via
`-p:CustomAfterMicrosoftCommonTargets=...` and
`-p:CustomAfterMicrosoftCommonCrossTargetingTargets=...`) while packing
`Ofdrw.Net.Reader` with the installed .NET 10 SDK. The nuspec then read:

```xml
<dependency id="Ofdrw.Net.Core" version="[0.3.1]" exclude="Build,Analyzers" />
<dependency id="Ofdrw.Net.Packaging" version="[0.3.1]" exclude="Build,Analyzers" />
<dependency id="JBig2Decoder.NETStandard" version="1.5.2" exclude="Build,Analyzers" />
```

Third-party dependencies were unchanged. It was **not** tried from a
`Directory.Build.targets` file, and no consumer restore was run against pinned
packages — Steps 2 and 5 do that.

`_GetProjectReferenceVersions` and `_ProjectReferencesWithVersions` are internal
names of the NuGet pack targets. They have been stable for years, but they are
not a public contract — hence the validator check in Step 3, which fails the
build loudly if a future SDK stops honouring the target.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| Script tests | `python3 -m unittest discover -s scripts/tests -v` | `OK` |
| Pack one | `dotnet pack src/Ofdrw.Net.Reader/Ofdrw.Net.Reader.csproj -c Release --no-build --no-restore -o "$OUT" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Successfully created package` |
| Show dependencies | `unzip -p "$OUT"/Ofdrw.Net.Reader.*.nupkg '*.nuspec' \| grep dependency` | (inspection) |
| Package E2E | `./scripts/run-converter-package-e2e.sh 0.3.1-plan029.1` | exit 0 |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

`$OUT` must be a directory outside the repository, e.g. `OUT="$(mktemp -d)"`.
The package E2E writes under `artifacts/` (git-ignored) and needs network access
to nuget.org for third-party packages unless the local cache already has them.

## Scope

**In scope** (the only files you should modify):

- `Directory.Build.targets` (create, repo root)
- `scripts/verify-package-artifacts.py`
- `scripts/tests/test_package_artifacts.py`
- `src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj` (remove one duplicate line)
- `docs/engineering/release-gates.md` and
  `docs/content/getting-started/quickstart.md` (one sentence each)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- Third-party `PackageReference` versions and
  `docs/engineering/baselines/third-party-dependencies.json`.
- Removing or narrowing `InternalsVisibleTo` grants (other than the duplicate).
- `Directory.Build.props` version values; workflows; the E2E scripts.

## Git workflow

- Branch: `codex/exact-sibling-package-versions`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Make the validator require exact ranges (test first)

In `scripts/tests/test_package_artifacts.py` add two tests next to
`test_wrong_sdk_dependency_is_rejected`, reusing its rewrite technique (extract
it into a small helper method `_with_core_dependency(self, version_text)` if
that keeps the tests short):

- `test_minimum_version_sdk_dependency_is_rejected` — dependency
  `id="Ofdrw.Net.Reader" version="{self.version}"` → `ValueError`.
- `test_exact_version_sdk_dependency_is_accepted` — dependency
  `version="[{self.version}]"` → `inspect` returns a manifest with 12 packages.

**Verify**: `python3 -m unittest discover -s scripts/tests -v` → the
"minimum version" test FAILS (no `ValueError` raised), the other passes.

Then change `scripts/verify-package-artifacts.py` line 53 to:

```python
                    if dependency.get('version') != f'[{version}]':
                        raise ValueError(f'SDK dependency must be the exact range [{version}] in {path.name}: {dependency.attrib}')
```

**Verify**: `python3 -m unittest discover -s scripts/tests -v` → `OK`.

### Step 2: Emit exact ranges at pack time

Create `Directory.Build.targets` in the repository root:

```xml
<Project>
  <!-- Ofdrw.Net packages use each other's internals (InternalsVisibleTo), so a package only works
       with siblings of the same build. Pack sibling ProjectReferences as exact ranges [x.y.z].
       _GetProjectReferenceVersions/_ProjectReferencesWithVersions are NuGet pack internals;
       scripts/verify-package-artifacts.py fails if they ever stop producing exact ranges. -->
  <Target Name="OfdrwExactSiblingPackageVersions" AfterTargets="_GetProjectReferenceVersions"
          Condition="'$(IsPackable)' != 'false'">
    <ItemGroup>
      <_OfdrwExactReference Include="@(_ProjectReferencesWithVersions)">
        <ProjectVersion>[%(_ProjectReferencesWithVersions.ProjectVersion)]</ProjectVersion>
      </_OfdrwExactReference>
      <_ProjectReferencesWithVersions Remove="@(_ProjectReferencesWithVersions)" />
      <_ProjectReferencesWithVersions Include="@(_OfdrwExactReference)" />
    </ItemGroup>
  </Target>
</Project>
```

**Verify**: build → `0 Error(s)` with the same warning count as before the file
existed (record both). Then run "Pack one" and "Show dependencies": both
`Ofdrw.Net.Core` and `Ofdrw.Net.Packaging` lines show `version="[0.3.1]"` in
both target-framework groups, and third-party lines are unchanged
(`JBig2Decoder.NETStandard` `1.5.2`, `SixLabors.ImageSharp` `2.1.13`).

### Step 3: Prove the whole default set and a real consumer

Run `./scripts/run-converter-package-e2e.sh 0.3.1-plan029.1`. It builds, packs
the twelve default packages, runs the (now strict) validator, restores a clean
consumer from those packages only, installs the CLI tool from them and runs the
harness.

**Verify**: exit 0. Then confirm no package escaped:

```sh
for f in artifacts/package-e2e/0.3.1-plan029.1/packages/*.nupkg; do
  unzip -p "$f" '*.nuspec' | grep -o '<dependency id="Ofdrw\.Net\.[^"]*" version="[^"]*"' | grep -v 'version="\[' && echo "NOT EXACT: $f"
done
```

→ prints nothing.

### Step 4: Remove the duplicate grant

In `src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj` delete the second
`<InternalsVisibleTo Include="Ofdrw.Net.Converter.Pdf" />` (in the second
`ItemGroup`).

**Verify**: `grep -c 'InternalsVisibleTo Include="Ofdrw.Net.Converter.Pdf"' src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj`
→ `1`; build → `0 Error(s)`.

### Step 5: Optional packages

The optional packages (`Ofdrw.Net.Graphics.SkiaSharp`,
`Ofdrw.Net.Converter.Pdf.Vector`, `Ofdrw.Net.Signatures.SesInterop`, the crypto
packages) are packed by their own scripts and now also pin their siblings. Run
the two that CI runs against the candidate feed, reusing the packages from Step 3:

```sh
./scripts/run-skia-package-e2e.sh 0.3.1-plan029.1 artifacts/package-e2e/0.3.1-plan029.1/packages artifacts/plan029/skia
./scripts/run-pdf-vector-package-e2e.sh 0.3.1-plan029.1 artifacts/package-e2e/0.3.1-plan029.1/packages artifacts/plan029/pdf-vector
```

**Verify**: both exit 0. If one cannot run in your environment (missing native
dependency, no network), record the exact error and report the step as "not
verified locally — covered by CI job `package-e2e`"; do not skip silently.

### Step 6: Document

- `docs/engineering/release-gates.md`: add one sentence in the section that
  describes package validation (search for `package` / `manifest`): sibling
  `Ofdrw.Net.*` dependencies are packed as exact ranges by
  `Directory.Build.targets`, and `scripts/verify-package-artifacts.py` rejects
  anything else.
- `docs/content/getting-started/quickstart.md`, after the second
  `dotnet add package` block (the one installing `Ofdrw.Net.Layout` and
  `Ofdrw.Net.Packaging`), add one sentence in Chinese:
  `所有 Ofdrw.Net.* 包必须使用同一版本；从下一版本起，包之间的依赖声明为精确版本，混用版本会在还原时报错（main，尚未发布）。`

**Verify**: `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check`
→ all exit 0.

## Test plan

- Two new unit tests in `scripts/tests/test_package_artifacts.py` (Step 1).
- End-to-end proof by the existing package consumer script (Step 3).

## Done criteria

- [ ] `python3 -m unittest discover -s scripts/tests -v` → `OK`
- [ ] Build prints `0 Error(s)`
- [ ] `./scripts/run-converter-package-e2e.sh 0.3.1-plan029.1` exits 0 and the Step 3 loop prints nothing
- [ ] Step 5 results recorded (pass, or "not verified locally" with the error)
- [ ] `git status --short` lists only in-scope files (nothing under `artifacts/`); storage gate exits 0
- [ ] `plans/plan-index.md` status row updated

## STOP conditions

Stop and report back (do not improvise) if:

- After Step 2 the nuspec still shows `version="0.3.1"` for siblings — the
  internal target names may differ in your SDK. Report `dotnet --version`; do
  not rewrite nuspecs by post-processing the `.nupkg`.
- The consumer restore in Step 3 fails with `NU1107`, `NU1608` or a downgrade
  error — some project or harness mixes sibling versions on purpose. Report the
  full message.
- The build's warning count changes, or any project that is not packable starts
  failing.
- You find a package-referenced harness under `e2e/` that deliberately combines
  different `Ofdrw.Net.*` versions.

## Maintenance notes

- Consumers can no longer upgrade one `Ofdrw.Net.*` package alone. This must be
  called out under compatibility in the next release note.
- If a package is ever meant to float independently of the rest, it must first
  stop using sibling internals; only then exempt it from the target.
- The remaining `InternalsVisibleTo` surface (about twenty grants between
  shipping assemblies) is the underlying coupling; reducing it is separate,
  larger work and was deliberately not planned here.
