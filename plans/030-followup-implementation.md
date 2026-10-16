# 030 follow-up: implemented local archival preflight

- Frozen integration baseline: `3247c76f95c9a3312cc26144e24c4984767af26f` (`codex/audit-integration`). Implementation commit: `8d67ba3aeaffd60cc5a44d3d2f94862cbf153047`.
- Worktree: `/Users/wanghongyi/.codex/worktrees/audit-followup-030/ofdrw.net`; branch: `codex/audit-followup-030`.
- The follow-up explicitly authorizes implementation and local commits, superseding the historical design-only restriction in [the original spike](030-archival-preflight-spike.md) and [design](030-archival-preflight-design.md). Those historical design records remain unchanged. No push, PR, integration merge, tag or publication is performed here.
- Status: implemented and verified; ready for parent integration review. The final gates below ran against implementation commit `8d67ba3aeaffd60cc5a44d3d2f94862cbf153047`; subsequent changes only record this report. `ready_to_merge=true` means this scoped task is ready, not that integration, publication or certification has occurred.

## Delivered scope and drift

The current loader still loads an independent in-memory archive, retains borrowed input, validates ZIP budgets and rejects unsafe/duplicate entry names. Core still grants Packaging access to the internal package path resolver. The existing ValidationReport and structure checker retain their existing contracts; the implementation adds an independent `Ofdrw.Net.Packaging.Preflight` namespace, without Reader, converter, crypto, signatures, networking or serializer dependencies. The frozen source's loader/path/public types were re-read rather than treating the older design baseline as current proof.

[Usage and precise rule scope](../docs/archival-preflight.md) documents `ArchivalPreflightChecker.CheckAsync`, snapshotted policies/budgets, six APF rules, public four-state report/evidence types and deterministic result ordering. Run completeness is independent of rule outcomes. Truncation preserves earlier observations and adds six Unknown diagnostics using reserved output slots; all inapplicable stays NotApplicable. Known failure dominates Unknown while preserving both observations. Standard clauses remain empty and unverified; no certification field claims success.

Known control fields use owning namespaces and explicit containing-entry semantics. URI classification precedes package path resolution; root-relative references remain inside the ZIP. Unsupported namespaces/shapes/attribute ownership, annotation/composite scopes and conventional orphan resources remain Unknown with coverage gaps. Font and signature sentinels prove only nonempty/present payloads. Existing checker behavior is unchanged and there is no implicit lossy boolean adapter or new product CLI.

## Verification

Artifacts are local and ignored by Git: `/Users/wanghongyi/.codex/worktrees/audit-followup-030/ofdrw.net/artifacts/followup-030/`. Final source manifest/environment and final gate logs live under its `final/` directory. A report-only documentation commit may follow the implementation commit; executable source is pinned by the manifest.

- Targeted final source: **81/81 passed, zero failed/skipped**, `tests-namespace-final/preflight.trx`. Includes the 6×4 rule matrix, failure plus unknown aggregation, Warning failures, all-inapplicable/no-object behavior, malformed/unsafe/duplicate ZIP, DTD, unsupported encoding/namespaces, conflicting font and attachment data, missing/empty payloads, root-relative/BaseLoc paths, orphan entries, unsupported transforms, cycles, XML/input/reference/result budgets, cancellation, original I/O propagation, input bytes/position/lifetime, option snapshots and stable evidence.
- NoFetch: live loopback TCP listener remained without pending connections for resource URLs and external DTD; file URI/DTD sentinels were opened exclusively and remained unchanged. The checker never opens host files or instantiates a network client/resolver; a source dependency review corroborates these observations. This is bounded test evidence, not OS-wide syscall tracing.
- Final-source Packaging full regression: **1,208/1,208 passed, zero failed/skipped**, `final/packaging/packaging.trx`.
- Final-source solution regression: **2,893/2,893 passed across 11 TRX suites, zero failed/skipped**, `final/solution/*.trx`. This includes the 81 new cases; counts from separate runs are not additive.
- Final-source public API documentation Rebuild: **passed**, zero new CS1591 diagnostics, 272 historical symbols, one existing partial-type diagnostic cleared; no baseline edit. `final/docs.log`. Packaging compiles for both netstandard2.0 and netstandard2.1.
- Final-source local NuGet consumer: **4/4 passed**, `final/consumer-run.log` and `final/consumer-output/*.json`. Fresh cache and feed contain only `Ofdrw.Net.Core` and `Ofdrw.Net.Packaging` at `0.1.0-audit030.20261016.2`; project.assets.json confirms exactly those two package dependencies. Both consumed DLL hashes match the package contents. Inputs preserve their before/after SHA-256.

Full gates use the atomic shared `validation.lock` with task/thread/time owner metadata and `try/finally` release. No lease is removed or inferred stale. No GUI lease is acquired. All dotnet commands use a writable task home, first-time/telemetry opt-outs, and single-node/no-server/no-shared-compilation flags. Repository-pinned offline trusted fonts are prepared and required for the full solution run; TRX must have nonzero total and zero skips. The exact commands are persisted in `final/validation-commands.json` by the local orchestration script.

The package-only consumer is [this standalone E2E](../e2e/Ofdrw.Net.Packaging.Preflight.E2E/Program.cs), copied into an isolated artifacts consumer directory, restored from this run's Core/Packaging feed into a fresh package cache, with no project references. It exercises presence-only Pass (deliberately invalid font/signature bytes), missing font Fail, encrypted package Fail with obscured scopes Unknown, and all-inapplicable NotApplicable. It persists ZIP fixtures, JSON reports, unchanged before/after input SHA-256, runtime and consumed assembly SHA-256. Final package/consumer hashes are recorded in `final/validation-summary.json`. The input ZIP sizes are 2,018 bytes (presence-only), 1,900 (missing-font), 2,031 (encrypted) and 366 (all-inapplicable); preflight changes none of them.

## Artifact identity and environment

Final test/consumer environment: macOS 26.6.2 arm64; SDK 10.0.401; runtime .NET 10.0.12. This is compilation evidence for both library target frameworks and execution evidence on .NET10, not proof of runtime behavior on every .NET implementation. Trusted font bundle SHA-256 is `40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59`; no optional trusted-font tests were skipped.

| Final artifact | SHA-256 |
| --- | --- |
| Core nupkg `0.1.0-audit030.20261016.2` | `2cd4a114657b34a1e8877e68e1bf7caa958a00582dbcd25712b03855aa5724f8` |
| Packaging nupkg `0.1.0-audit030.20261016.2` | `599612df2179e0f63f6e0747e2552ff82c460bf02faec7f7cd9b77f3bd41aba7` |
| Packaged and consumed Packaging netstandard2.1 DLL | `b404ccaf8f159c7e1a3b7f01a6fc8b0057f713b60e07f5ac1f8bc8c092536f06` |
| Packaged and consumed Core netstandard2.1 DLL | `ee281d02d52846920613ddfdc508285b58d7d5cd71366910a1a26bf363c290ff` |

The source manifest lists every implementation file's SHA-256 and full code commit. `final/io-surface-review.json` records the I/O boundary inspection at the same checker hash. Earlier development logs outside `final/` remain available but are not substituted for these final gate results. The report links and source manifest were independently checked, and `git diff --check` passes.

## Visual acceptance

- Functional tests: **passed**, targeted, Packaging and solution results above.
- Automated page rendering: **not applicable**.
- PNG visual review: **not applicable**.
- macOS Preview: **not applicable**.

This is a read-only ZIP/XML inspection module with synthetic presence fixtures; it changes no conversion, layout, font rendering or generated page behavior. The four-byte font and invalid SignedValue are never rendered or used as validity evidence. Passing these checks does not establish appearance or reader compatibility.

## Deliberate coverage limits and handoff

No verified standard text/clause mapping or customer acceptance policy was supplied. The profile is local and experimental. Font validity/glyphs/licensing, decryption, hashes/signature trust, arbitrary XML/actions/payload internals, full ID closure, XSD validation, conformance/certification, real-world corpus precision and peak-memory/latency benchmarking remain outside this first version. The public loader's existing central-directory allocation behavior is inherited, not upgraded to the .NET10-only accounted loader's stronger profile. Conservative Unknown results are expected for unmodeled scopes.

The separate Graphics/Vector optional project suites and page/Preview gates are not claimed by this Packaging-only implementation. Parent coordination owns five-task integration, combined-head verification and any eventual merge to the integration branch. This task never merges or releases by itself.
