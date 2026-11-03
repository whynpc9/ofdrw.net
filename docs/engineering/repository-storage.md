# Repository storage

## Default checkout

Binary archives, document/image fixtures and fonts are stored in Git LFS.
`.lfsconfig` downloads only `e2e/**` and `tests/**` by default. CI skips automatic
smudging and explicitly fetches those inputs; the public-table evidence job
also fetches its own retained bundle. This avoids downloading all historic
visual evidence on every build.

Use the [clone instructions](../../README.md#clone-and-large-files) for a shallow
checkout before the history cutover, or to fetch one evidence directory.
`git archive` exports LFS pointers rather than hydrated payloads: use an LFS
checkout when running tests or reviewing binary inputs from a source snapshot.

## What to retain

- Keep deterministic, shared input fixtures and regression tests that exercise
  distinct behavior. Store binary fixtures in LFS, and reuse them across tests.
- Keep a compact report of the source commit, environment, commands, actual
  results, inspected pages and remaining limitations.
- Keep only the input/output pages, screenshots or failed-run records required
  to reproduce a defect or assess the reported acceptance scope. Put large
  retained payloads in one optional LFS bundle with a size/SHA-256 manifest.
- Leave build output, caches, repeated source trees and routine full test logs
  in ignored `artifacts/` or CI artifacts. A new source commit is not a reason
  to retain another copy of every previous artifact.

Existing validation history is preserved. The 1,171 directly tracked PR23
artifact files were consolidated without changing their bytes: 42.4 MiB became
one 7.6 MiB compressed LFS bundle. Its [index and manifest](../evidence/pr23-integration/README.md)
restore the original paths. No existing .NET regression tests were removed;
the large footprint came from retained evidence rather than test source.
Routine solution-test console verbosity was reduced while retaining the tests
and existing TRX reporting.

Run `python3 scripts/check-repository-storage.py` after staging changes. CI
checks the staged/checked-out Git objects, rejects tracked `artifacts/` files
and ordinary blobs above 5 MiB, and verifies required fixture hashes. This
check works with optional evidence left as pointer files.

## Historical migration

历史迁移演练和切换方案已从现行操作说明移除。需要追溯时阅读[迁移前记录](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/repository-storage.md#history-cutover)。它描述原时间点的操作，不是当前待办或改写远端历史的授权。
