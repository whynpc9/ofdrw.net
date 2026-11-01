# Audit025 allocation measurement boundary

The original F2 main Windows run measured 7,760 bytes for one 4,096-character
whitespace-delta call. That first failure remains recorded; a later passing run
does not explain it. No Core coverage, positioning, font, or conversion code is
changed by this test-harness correction.

`Audit025_ElementOverloadNoDelta` now invokes its own test assembly as an isolated
child. Microsoft.NET.Test.Sdk already builds an executable; its generated empty
entry point is replaced with an explicit diagnostic entry point. Ordinary
`dotnet test` continues to discover and execute the test assembly through VSTest.
An invocation without arguments remains inert; unsupported or incomplete
arguments fail with exit 2.

The budget probe checks null, empty, and whitespace deltas at 0, 32, 4,096, and
8,192 characters. Each interval has exactly one recorded warmup and eight
recorded measured calls. Every measured target call must allocate less than the
original 4,096-byte limit. Empty, intentional 8,192-byte allocation, and former
per-grapheme implementations verify measurement sensitivity. A separate
regression-control child routes the target through that former implementation;
the same budget evaluator must reject every long run, with exit 1.

Inputs and result storage are allocated before each bounded 8 MiB no-GC region.
Region startup, latency mode, unchanged collection counts, measurement thread,
and successful region end are required. An unavailable or invalidated region is
a failure, with no retry or fallback. No minimum, average, threshold increase,
or skipped assertion is used. Parent processes have a 20-second limit and kill
their owned child tree on timeout. Every receipt uses non-overwriting creation
and records the actual Core path/hash/target framework/version, fixture hash,
runtime, architecture, GC configuration, interval status, and raw counters.

CI also runs one separately named, bounded background-GC pressure diagnostic
and retains all 128 intervals, including empty, actual target, positive, and
per-grapheme controls. Positive deltas in this diagnostic are observations, not
budget-test retries. See the upstream [counter report](https://github.com/dotnet/runtime/issues/134724)
and [draft accounting change](https://github.com/dotnet/runtime/pull/134855) for
one hypothesis; neither identifies the cause of the original Windows call.
Collection-count changes alone do not attribute an individual delta.

After a Release build, run `python3 scripts/run-audit025-diagnostic.py`. Raw
receipts are retained under `artifacts/audit025-diagnostics`, or the explicit
`OFDRW_AUDIT025_DIAGNOSTICS_DIRECTORY`, and uploaded with full-suite CI results.
No visual capability, original Skia failure, quantity-counter defect, or native
Preview Find gate is closed by this test-only work.
