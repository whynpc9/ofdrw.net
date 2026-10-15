# Password envelope fixture and actual package consumer

The project generates deterministic licensed native/default DOCX→OFD fixtures,
then tests selected second-page and all-entry recovery. It compares every original
ZIP file entry's bytes and exports recovered OFD through the existing OFD→PDF
converter. The Noto font is pinned to SHA-256 `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa`
and OFL license accompanies the output. Its large embedded font requires explicitly
raised 32 MiB entry / 128 MiB expanded / 160 MiB output fixture budgets.

`../../scripts/run-password-crypto-e2e.sh` packs Core, Packaging and the independent
password extension into a local feed and creates a clean PackageReference-only
consumer with an isolated cache and product source mapping. Consumer.cs.txt uses
both generated baseline OFDs, verifies all four actual-package recoveries by exact
entry name/bytes, and verifies wrong-password/cancellation preserving existing output.
It checks the v2 descriptor and distinct salts, and rejects tag tampering, excessive
KDF parameters and v1 confusion while preserving the restored output.
No ProjectReference or installed Ofdrw.Net cache can satisfy that consumer.

```sh
export DOTNET_CLI_HOME=/tmp/ofdrw-password-e2e-home DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
python3 scripts/install-ci-fonts.py --directory artifacts/password-crypto/fonts
dotnet build e2e/Ofdrw.Net.Crypto.Password.E2E -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet e2e/Ofdrw.Net.Crypto.Password.E2E/bin/Release/net10.0/Ofdrw.Net.Crypto.Password.E2E.dll "$PWD" "$PWD/artifacts/password-crypto/current" "$PWD/artifacts/password-crypto/fonts"
./scripts/run-password-crypto-e2e.sh
```

Profile v2 authenticates logical file names and bytes with HMAC before decrypting;
its recovery checksums provide additional fault/structure checks. It rejects v1
without fallback. This is not a GM/T integrity protocol implementation. Tests do
not establish vendor interoperation, certification or archival acceptance. Actual
v2 Preview scope must be recorded separately; the initial/review1/review2/review3
records describe only historical v1. Automated PDF export is not Preview.
