# Plan 032: Signature verification resolves package references like the rest of the SDK

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Core/IO/OfdPackagePath.cs src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs`
> If either file changed since this plan was written, compare the "Current
> state" excerpts against the live code before proceeding; on a mismatch, treat
> it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: MED (signature verification; every change must fail closed)
- **Depends on**: none
- **Category**: bug, tech-debt
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

References inside an OFD package (`BaseLoc`, `SignedValue`, `FileRef`, …) are
paths relative to the file that contains them. The SDK has one shared resolver,
`OfdPackagePath.Resolve`, used by about 34 call sites: it drops `.` segments,
applies `..`, and rejects a path that leaves the package root. Six other
classes carry their own private copy, each with different rules.

The copy in the signature verifier does not apply `.` or `..` at all. Checked
at `427433a` with a package whose reference digests are correct
(`ofdrw verify-signatures`):

| Reference in the package | Result |
|---|---|
| `BaseLoc="Signature.xml"` | exit 2 — reference integrity valid |
| `BaseLoc="./Signature.xml"` | exit 3 — `Signature description is missing: Signs/./Signature.xml` |
| `BaseLoc="../Signs/Signature.xml"` | exit 3 — `Signature description is missing: Signs/../Signs/Signature.xml` |
| `<SignedValue>../Signs/SignedValue.dat</SignedValue>` | exit 3 — `Signed value is missing: Signs/../Signs/SignedValue.dat` |
| `FileRef="/./content.txt"` | exit 3 — references 0/1 |

All five are the same valid package; the reader and the stamp renderer resolve
those references without difficulty. The verifier fails closed, so this is not
a way to forge a result — it is a correctly signed document from another
producer being reported as invalid, and a second definition of "which file does
this reference mean" inside the component where that question matters most.

After this plan the verifier uses the shared resolver. References that leave
the package root are still errors.

## Current state

### The shared resolver

`src/Ofdrw.Net.Core/IO/OfdPackagePath.cs` (complete relevant part):

```csharp
internal static class OfdPackagePath
{
    internal static string Resolve(string containingEntry, string reference)
    {
        var value = reference.Replace('\\', '/');
        var combined = value.StartsWith("/", StringComparison.Ordinal)
            ? value : GetDirectory(containingEntry) + "/" + value;
        var parts = new List<string>();
        foreach (var part in combined.Split('/'))
        {
            if (part.Length == 0 || part == ".") continue;
            if (part == "..")
            {
                if (parts.Count == 0) throw new InvalidDataException("Package reference escapes the OFD root.");
                parts.RemoveAt(parts.Count - 1);
            }
            else parts.Add(part);
        }

        return string.Join("/", parts);
    }

    internal static string GetDirectory(string entry)
    {
        var index = entry.LastIndexOf('/');
        return index < 0 ? string.Empty : entry.Substring(0, index);
    }
```

`Ofdrw.Net.Core` grants `InternalsVisibleTo` to `Ofdrw.Net.Signatures` and to
`Ofdrw.Net.Core.Tests`. There are no unit tests for `OfdPackagePath` today.

### The verifier's private copy

`src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs:423-444`:

```csharp
    private static string ResolvePath(string containingPath, string reference)
    {
        if (reference.StartsWith("/", StringComparison.Ordinal))
        {
            return NormalizePath(reference);
        }

        var slash = containingPath.LastIndexOf('/');
        var directory = slash < 0 ? string.Empty : containingPath.Substring(0, slash + 1);
        return NormalizePath(directory + reference);
    }

    private static string NormalizePath(string path)
    {
        return path.Replace('\\', '/').TrimStart('/');
    }

    private static bool TryNormalizePath(string path, out string normalized)
    {
        normalized = NormalizePath(path);
        return !normalized.Split('/').Any(segment => segment == "..");
    }
```

Its three uses on references:

Signature description (line 183):

```csharp
            var signaturePath = ResolvePath(listPath, baseLocation!);
            result.SignaturePath = signaturePath;
            if (!archive.TryGetBytes(signaturePath, out var signatureBytes))
            {
                result.CryptographicStatus = OfdCryptographicVerificationStatus.Missing;
                result.CryptographicMessage =
                    $"Signature description is missing: {signaturePath}";
                continue;
            }
```

Protected file references (lines 281–299):

```csharp
            var fileReference = reference.Attribute("FileRef")?.Value ?? string.Empty;
            var referenceResult = new OfdReferenceVerificationResult
            {
                FileReference = fileReference,
                ResolvedEntryName = NormalizePath(fileReference),
                DigestAlgorithmSupported =
                    OfdDigestAlgorithms.IsSupported(result.CheckMethod)
            };
            result.References.Add(referenceResult);
            if (!TryNormalizePath(fileReference, out var entryName))
            {
                referenceResult.Error = "The reference path is unsafe.";
                continue;
            }

            referenceResult.ResolvedEntryName = entryName;
            if (!archive.TryGetBytes(entryName, out var entryBytes))
            {
                referenceResult.Error = "The referenced package entry is missing.";
                continue;
            }
```

Signed value (line 359):

```csharp
        var signedValuePath = ResolvePath(signaturePath, signedValueLocation!);
        result.SignedValuePath = signedValuePath;
        if (!archive.TryGetBytes(signedValuePath, out var signedValue))
        {
            result.CryptographicStatus = OfdCryptographicVerificationStatus.Missing;
            result.CryptographicMessage =
                $"Signed value is missing: {signedValuePath}";
            return;
        }
```

The status enum
(`src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationModels.cs:6-13`):
`Missing, Unsupported, Valid, Invalid, Error`. A signature whose status is
`Error` or `Missing` is never "fully valid".

Line 140 (`var listPath = NormalizePath(location);`, the signatures list named
in `OFD.xml`) is out of scope: it is not resolved against a containing file.

### A fixture that needs no real certificate

`tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs` (`WriteDeclarationsFixture`,
about line 161) builds a package whose signature uses an unregistered algorithm,
so verification stops at "reference integrity valid" without any key material:

```csharp
        XNamespace ns = "http://www.ofdspec.org/2016";
        using (var zip = ZipFile.Open(input, ZipArchiveMode.Create))
        {
            var body = new XElement(ns + "DocBody", new XElement(ns + "Signatures", "Signs/Signatures.xml"));
            WriteVerificationEntry(zip, "OFD.xml", new XElement(ns + "OFD", body).ToString());
            WriteVerificationEntry(zip, "content.txt", "protected content");
            WriteVerificationEntry(zip, "Signs/Signatures.xml", new XElement(ns + "Signatures", new XElement(ns + "Signature",
                new XAttribute("ID", "1"), new XAttribute("BaseLoc", "Signature.xml"))).ToString());
            var checkValue = Convert.ToBase64String(OfdDigestAlgorithms.Compute(
                OfdDigestAlgorithms.Sha256Oid, Encoding.UTF8.GetBytes("protected content")));
            var references = new XElement(ns + "References",
                new XAttribute("CheckMethod", OfdDigestAlgorithms.Sha256Oid),
                new XElement(ns + "Reference", new XAttribute("FileRef", "/content.txt"),
                    new XElement(ns + "CheckValue", checkValue)));
            var signature = new XElement(ns + "Signature",
                new XElement(ns + "SignedInfo",
                    new XElement(ns + "SignatureMethod", "urn:audit001:unregistered"), references),
                new XElement(ns + "SignedValue", "SignedValue.dat"));
            WriteVerificationEntry(zip, "Signs/Signature.xml", signature.ToString());
            WriteVerificationEntry(zip, "Signs/SignedValue.dat", "unregistered algorithm fixture");
        }
```

and it is checked like this:

```csharp
            var report = await new OfdSignatureVerifier().VerifyAsync(stream);
            Assert.True(Assert.Single(report.Signatures).ReferenceIntegrityValid);
            Assert.Equal(OfdCryptographicVerificationStatus.Unsupported, report.Signatures[0].CryptographicStatus);
```

`tests/Ofdrw.Net.Signatures.Tests/Audit001RegressionTests.cs` contains a similar
in-memory builder (`CreatePackage`, a `ZipArchive` over a `MemoryStream`) — use
it as the structural pattern for the new test file.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| Core tests | `dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Signature tests | `dotnet test tests/Ofdrw.Net.Signatures.Tests/Ofdrw.Net.Signatures.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| CLI tests | `dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| SES interop tests | `dotnet test tests/Ofdrw.Net.Signatures.SesInterop.Tests/Ofdrw.Net.Signatures.SesInterop.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Core/IO/OfdPackagePath.cs`
- `src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs`
- `tests/Ofdrw.Net.Core.Tests/OfdPackagePathTests.cs` (create)
- `tests/Ofdrw.Net.Signatures.Tests/ReferencePathResolutionTests.cs` (create)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch) — the other private resolvers. They are listed
so that nobody "finishes the job" inside this plan; each needs its own
behaviour study:

| Location | Current rule |
|---|---|
| `src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs:468-508` | applies `.`/`..`; a `..` above the root is silently dropped; whitespace-only segments dropped |
| `src/Ofdrw.Net.Packaging/Validation/OfdPackageStructureChecker.cs:153-189` | same as the appearance reader |
| `src/Ofdrw.Net.Signatures/Signing/OfdSignatureFileSelection.cs:327-345` | applies `.`/`..`; a `..` above the root yields an empty path |
| `src/Ofdrw.Net.Signatures/Signing/OfdSignatureService.cs:472-488` | throws on any `..` segment |
| `src/Ofdrw.Net.Reader/Readers/OfdReader.cs:1668` (`GetDirectory`) | directory of a normalized path |
| `src/Ofdrw.Net.Packaging/Archive/OfdPackageArchive.cs:50` (`Normalize`) | slash and leading-slash normalization only |

Also out of scope: digest computation, algorithm registration, exit codes, the
verification report's public shape.

## Git workflow

- Branch: `codex/signature-reference-resolution`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: A non-throwing form of the shared resolver, with tests

In `OfdPackagePath.cs` move the body of `Resolve` into a new method and let
`Resolve` call it. Behaviour of `Resolve` must not change.

```csharp
    internal static bool TryResolve(string containingEntry, string reference, out string resolved)
    {
        var value = reference.Replace('\\', '/');
        var combined = value.StartsWith("/", StringComparison.Ordinal)
            ? value : GetDirectory(containingEntry) + "/" + value;
        var parts = new List<string>();
        foreach (var part in combined.Split('/'))
        {
            if (part.Length == 0 || part == ".") continue;
            if (part == "..")
            {
                if (parts.Count == 0) { resolved = string.Empty; return false; }
                parts.RemoveAt(parts.Count - 1);
            }
            else parts.Add(part);
        }

        resolved = string.Join("/", parts);
        return true;
    }

    internal static string Resolve(string containingEntry, string reference) =>
        TryResolve(containingEntry, reference, out var resolved)
            ? resolved
            : throw new InvalidDataException("Package reference escapes the OFD root.");
```

Create `tests/Ofdrw.Net.Core.Tests/OfdPackagePathTests.cs` (namespace
`Ofdrw.Net.Core.Tests`, `using Ofdrw.Net.Core.IO;`) with a `[Theory]`:

| containingEntry | reference | resolved |
|---|---|---|
| `Doc_0/Signs/Signatures.xml` | `Sign_0/Signature.xml` | `Doc_0/Signs/Sign_0/Signature.xml` |
| `Doc_0/Signs/Signatures.xml` | `./Sign_0/Signature.xml` | `Doc_0/Signs/Sign_0/Signature.xml` |
| `Doc_0/Signs/Signatures.xml` | `../Res/a.png` | `Doc_0/Res/a.png` |
| `Doc_0/Signs/Signatures.xml` | `/Doc_0/./Pages//Page_0/Content.xml` | `Doc_0/Pages/Page_0/Content.xml` |
| `Doc_0/Signs/Signatures.xml` | `Sign_0\Signature.xml` | `Doc_0/Signs/Sign_0/Signature.xml` |
| `OFD.xml` | `Doc_0/Document.xml` | `Doc_0/Document.xml` |
| (empty string) | `content.txt` | `content.txt` |

Each row asserts both `TryResolve` (true + value) and `Resolve` (value). A
second theory for escapes — (`OFD.xml`, `../x`), (`Doc_0/a.xml`, `../../x`),
(`Doc_0/a.xml`, `/../x`) — asserts `TryResolve` is false and `Resolve` throws
`InvalidDataException`.

**Verify**: build → `0 Error(s)`; Core tests filtered with
`--filter "FullyQualifiedName~OfdPackagePathTests"` → all pass. Then run the
full Core, Packaging (`tests/Ofdrw.Net.Packaging.Tests`) and Reader-dependent
CLI test projects → `Failed: 0` (proves `Resolve` is unchanged).

### Step 2: Verifier tests (failing first)

Create `tests/Ofdrw.Net.Signatures.Tests/ReferencePathResolutionTests.cs`. Write
a private builder `MemoryStream Package(string baseLoc, string fileRef, string signedValue)`
that produces the fixture shown in "Current state" into a `MemoryStream`
(entries `OFD.xml`, `content.txt`, `Signs/Signatures.xml`, `Signs/Signature.xml`,
`Signs/SignedValue.dat`), with the three parameters substituted for
`"Signature.xml"`, `"/content.txt"` and `"SignedValue.dat"`.

`[Theory] EquivalentReferencesVerifyAlike` — each row must give
`ReferenceIntegrityValid == true` and
`CryptographicStatus == OfdCryptographicVerificationStatus.Unsupported`:

| baseLoc | fileRef | signedValue |
|---|---|---|
| `Signature.xml` | `/content.txt` | `SignedValue.dat` |
| `./Signature.xml` | `/content.txt` | `SignedValue.dat` |
| `../Signs/Signature.xml` | `/content.txt` | `SignedValue.dat` |
| `/Signs/Signature.xml` | `/content.txt` | `SignedValue.dat` |
| `Signature.xml` | `/./content.txt` | `SignedValue.dat` |
| `Signature.xml` | `content.txt` | `SignedValue.dat` |
| `Signature.xml` | `/content.txt` | `../Signs/SignedValue.dat` |
| `Signature.xml` | `/content.txt` | `./SignedValue.dat` |

`[Theory] ReferencesThatLeaveThePackageAreErrors`:

| baseLoc | fileRef | signedValue | Expect |
|---|---|---|---|
| `../../Signature.xml` | `/content.txt` | `SignedValue.dat` | status `Error`; `CryptographicMessage` contains `package root`; `report.FullyValid` false |
| `Signature.xml` | `/content.txt` | `../../SignedValue.dat` | status `Error`; message contains `package root` |
| `Signature.xml` | `/Signs/../content.txt` | `SignedValue.dat` | `References[0].Error == "The reference path is unsafe."`; `ReferenceIntegrityValid` false |
| `Signature.xml` | `/../content.txt` | `SignedValue.dat` | same as previous row |

**Verify**: Signature tests filtered with
`--filter "FullyQualifiedName~ReferencePathResolutionTests"` → rows 1, 4 and 6
of the first theory pass; rows 2, 3, 5, 7, 8 FAIL; in the second theory the two
`FileRef` rows pass and the two `Error` rows FAIL (status is `Missing` today).
If the pattern differs, STOP.

### Step 3: Use the shared resolver in the verifier

Add `using Ofdrw.Net.Core.IO;` if absent. Then:

1. Signature description (line 183):

```csharp
            if (!OfdPackagePath.TryResolve(listPath, baseLocation!, out var signaturePath))
            {
                result.CryptographicStatus = OfdCryptographicVerificationStatus.Error;
                result.CryptographicMessage = "Signature BaseLoc leaves the package root.";
                continue;
            }
            result.SignaturePath = signaturePath;
```

2. Signed value (line 359):

```csharp
        if (!OfdPackagePath.TryResolve(signaturePath, signedValueLocation!, out var signedValuePath))
        {
            result.CryptographicStatus = OfdCryptographicVerificationStatus.Error;
            result.CryptographicMessage = "SignedValue location leaves the package root.";
            return;
        }
        result.SignedValuePath = signedValuePath;
```

3. Protected file references: keep the existing verdict for any `..` segment
   (a `FileRef` is an absolute package path; `..` in it stays "unsafe"), and
   resolve everything else with the shared code:

```csharp
            if (!TryNormalizePath(fileReference, out _) ||
                !OfdPackagePath.TryResolve(string.Empty, fileReference, out var entryName))
            {
                referenceResult.Error = "The reference path is unsafe.";
                continue;
            }
```

   Leave the initial `ResolvedEntryName = NormalizePath(fileReference)` as it is.

4. Delete the private `ResolvePath` method if it has no remaining callers
   (`grep -n "ResolvePath(" src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs`).
   Keep `NormalizePath` and `TryNormalizePath`.

Do not change any message that an existing test asserts; the two new messages
are additions.

**Verify**: build → `0 Error(s)`; `ReferencePathResolutionTests` → all pass.

### Step 4: Regression

Run the full Signatures, SES interop, CLI and Core test projects.

**Verify**: all `Failed: 0`. Additionally confirm end to end that nothing which
verified before changed: the first row of `EquivalentReferencesVerifyAlike` is
the pre-existing behaviour and must still pass, and
`Audit001_CliIncompleteReport` in the CLI tests must still pass.

## Test plan

- `tests/Ofdrw.Net.Core.Tests/OfdPackagePathTests.cs`: 7 resolving rows, 3
  escaping rows.
- `tests/Ofdrw.Net.Signatures.Tests/ReferencePathResolutionTests.cs`: 8
  equivalent-reference rows, 4 error rows. Pattern:
  `tests/Ofdrw.Net.Signatures.Tests/Audit001RegressionTests.cs`.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] Both new test classes pass; Core, Signatures, SES interop and CLI test projects → `Failed: 0`
- [ ] `grep -n "static string ResolvePath" src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs` → no matches
- [ ] No file outside the in-scope list is modified (`git status --short`); storage gate exits 0
- [ ] `plans/plan-index.md` status row updated

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- The fail-before pattern in Step 2 differs from the one stated.
- Any existing signature, SES interop or CLI test changes its result after
  Step 3. Do not adjust existing expectations.
- The SES interop tests cannot run in your environment — report that; do not
  mark the plan done on the strength of the other suites alone.
- You find that a real signed sample in `tests/fixtures` or `e2e/` depends on
  the old "no `..` handling" behaviour.

## Maintenance notes

- This plan makes the verifier accept more spellings of the *same* entry; it
  never makes a reference point at a different entry than the reader would
  open. Reviewers should check exactly that property: each resolved name is
  what `OfdPackagePath.Resolve` returns for the same inputs.
- The out-of-scope table is the remaining consolidation work. The two that
  silently drop a `..` above the root (appearance reader, structure checker)
  are the next candidates; changing them alters which stamp image is drawn for
  malformed packages and therefore needs visual verification under the
  repository's rendering rules.
