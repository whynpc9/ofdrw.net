# Independent frozen v2 execution verification

Candidate: `d4eb787585c0a13a992c1ca2c6cefe5ed5e0cbdf` (`codex/gmt0099-password-crypto`). Source was extracted from the commit into the plugin-provided temporary staging root. All 531 selected tracked source, test, E2E, script, required document and configuration blobs matched the frozen commit before and after execution. No repository source was edited.

## Executed gates

- Password security/project tests: 98 passed, 0 failed, 0 skipped. Includes independent scrypt and MAC vectors; v2 salt/IV randomness; strict KDF cost, version and descriptor rejection; authenticated metadata, entry/path and tag tampering; cancellation and atomic output.
- Full `dotnet test Ofdrw.Net.sln -c Release`: 2,309 passed, 0 failed, 12 skipped of 2,321. All 12 skips were ordinary real-font tests without the mandatory font environment.
- Forced original-font gate `scripts/run-trusted-font-tests.py`: 65 passed, 0 failed, 0 skipped. This separately exercises the required real-font cases.
- Python `scripts/tests`: 20 passed.
- Fresh optional package consumer: packed exactly Core, Packaging, Crypto.Password version `0.1.0-issue11.v2freeze`; source-mapped, isolated NuGet consumer has only one direct `Ofdrw.Net.Crypto.Password` PackageReference. Four native/default × partial/all cases recovered exact OFD entries, emitted distinct 32-byte v2 salts and authenticated envelopes, and preserved sentinel output on wrong password, cancellation, tampered tag, excessive cost, v1 descriptor, input-budget and raw-ZIP failures. The optional package's README bytes equal frozen `docs/password-crypto.md`; packaged DLL bytes equal the frozen build; BouncyCastle 2.6.2 occurs in the optional package dependency only.
- Fresh default package consumer: packed and verified exactly 11 default products at the same version. Actual tool install, CLI DOCX/OFD/image path, PDF converter consumer and font-subset consumer succeeded. The default feed contains no Crypto.Password package.

The ordinary full test run reported existing NuGet vulnerability warnings for SixLabors.ImageSharp 2.1.13. This validation did not change dependencies. No macOS Preview inspection was performed in this independent execution role; functional and package evidence is separate from visual acceptance.

## Retained evidence

- [Frozen source identity](frozen-source-check-final.json)
- [Full test counts](full-test-summary.json)
- [Forced font summary](trusted-fonts-summary.json)
- [Optional package identity, hashes and dependencies](password-package-identity.json)
- [Optional package functional cases](password-package-functional.json)
- [Optional package manifest](password-package-manifest.json)
- [Default package summary](default-package-summary.json) and [11-package manifest](default-package-manifest.json)

The temporary staging area holds the generated packages and consumer outputs. Retained records contain only public-compatible hashes, metadata, test counts and selected logs; NuGet caches, CLI home, source build trees, font binaries and commercial-font diagnostic output were not imported.
