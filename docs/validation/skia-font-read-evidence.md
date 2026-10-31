# Skia font-read evidence

This change adds first-failure evidence and earlier raw-style admission. It does
not fix the recurring rejection of supported font fixtures by the native backend.
It adds no retry, regular-face fallback, dependency change or test serialization.

The owned binding parses its bounded immutable clone's original OS/2/head
directory entries and flags before creating SKData/Typeface. This is a style
preflight, not full-font sanitization. Native table upper bounds, both native
style words, fixed-face checks, full original stream equality/SHA, original
glyph mapping and destination identity/atomicity remain required. The early
clone descriptor/hash never replaces native stream authentication. Malformed
input can now fail earlier with an original-style message; its exception type
remains NotSupportedException. Valid raw-style input with an unavailable native
table still fails. SKData retains its original constructor-wide lifetime.

Each existing size/one-byte read result is retained once and used for admission.
The internal exception-data key `Ofdrw.Skia.FontReadFailure.v1` carries bounded
structural metadata: clone length/SHA, raw offsets/lengths/flags, generation,
saved wrapper/handle, thread, phase, ordinal, monotonic start/completion times
and actual results. There is no font
payload, family/resource name or user text. Native handles are historical values;
diagnostics do not dereference them after failure. A guard's evidence is attached
before its first throw. A table call that itself throws retains its original
exception; later decoration is explicitly marked and is not original
first-chance evidence. Nothing is re-read or re-thrown solely to collect data.

Recording and attachment are best-effort and cannot skip cleanup or replace the
original failure. Per-invocation injection tests cover a first short size whose
hypothetical second query would succeed, a failed first byte, raw rejection
before native construction, lazy-batch atomicity, direct exceptions, failed
recording/attachment, and successful authentication/disposal with recording
unavailable. Existing style, same-name identity, original-byte and late-mismatch
tests remain unchanged.

With `OFDRW_SKIA_DIAGNOSTICS_DIRECTORY` configured, the Graphics/Vector test hosts
capture guard first-chance evidence and the same process's loaded-module paths
and file hashes, OS/RID and managed identities. A decorated propagation event is
labelled separately. macOS uses an existing-loader NOLOAD reference and dladdr;
the font symbol is never called and the temporary loader reference closes. No
second Skia library is loaded. Unsupported/ambiguous collection is recorded as
such, not inferred from expected NuGet files. Mapped-path hashes are disk files
at collection time, not memory hashes or independent per-call address attestation.
Uninjected guards without a raw clone preflight can be expected unsupported-input
rejections; they are labelled separately and are not valid-input recurrence proof.

Each binding keeps at most32 steps; optional process history keeps128. The test
host retains at most16 controlled and16 observed reports per process, with at
most32 deduplication entries. Capture failures never mask the source exception.
CI retains these files with the original TRX. Without optional collection, there
is no file/module inspection or process-history recording on the normal path.

A successful run or controlled injection does not establish that recurrence was
remediated. Actual failures remain visible and retain their first values. Fresh
immutable source/package/native-platform and actual-page visual gates remain
separate requirements.
