# PR 19 path-policy R2 independent affected gate

The exact `6b12e018c21f905904d1994b6c5b8f3875e1e1cf` Git archive was built in an isolated `source/` directory. All 387 archived source files matched HEAD Git blobs before and after; the tracked checkout remained clean. Commands, environment, raw logs, TRX and per-case results are preserved beside this report.

Fresh Release results: Core **69/69**, Cli **39/39**, and filtered Packaging **74/74**, with no failures or skips. The seven Packaging class counts were OwnedXmlPreservation 24, OwnedSaveBoundary 6, OwnedSaveObservation 2, OwnedPayloadClone 9, TextReplacementSpool 10, TextReplacementFileSession 16 and TextReplacementOwnedSession 7. Core, Reader, Layout and Converter.Pdf each built successfully for `netstandard2.0` and `netstandard2.1` with zero errors.

The two new Core cases and four new Cli theory rows executed and passed. The actual macOS 26.6.2 case-variant tests reported that `Input.ofd` and `input.ofd` aliased on this volume and exercised the conservative Windows/macOS rejection for both no-op and changed replacements; they do **not** demonstrate Linux independent-file publication. The valid-OFD same-canonical-path and actual symlink-to-input cases both executed and rejected while preserving files. The Windows symlink API permission/unavailability branch was not executed on this host. Existing ImageSharp audit/nullable warnings are retained in raw build logs.

This is the affected path-policy gate only; it does not repeat the full suite, mandatory 65-font gate, 11-package consumer, CLI success matrix or GUI visual acceptance.
