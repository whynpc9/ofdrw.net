# Immutable native asset delivery

`assets/manifest.json` binds the reviewed C source SHA-256 to exactly four
executables. A source change requires rebuilding the entire set. The former
`artifacts/pr23-process-reaper/native-prototype/` outputs are not inputs to this
build and must never be promoted.

Run on macOS with Docker and the exact CLT/SDK recorded below:

```text
python3 scripts/build-process-supervisor.py --source-sha256 <reviewed-source-sha256>
python3 scripts/verify-process-supervisor-assets.py --allow-unpublished
python3 scripts/verify-process-supervisor-assets.py --help
python3 scripts/build-process-supervisor.py --source-sha256 <reviewed-source-sha256> --promote-execution-proof <six-environment-records.json>
```

The builder downloads only the official musl source archive and verifies its
SHA-256; `--musl-source <cached-tarball>` supplies the same verified bytes offline.
Docker images must already be available by their immutable digests or Docker may
pull them. Each Linux build container itself has `--network=none`. The builder
compiles the executable twice, compares bytes, and validates the whole staged set
before replacing committed assets. A build emits a candidate manifest with
`publishable: false`. Successful compilation and matching hashes do not establish
runtime acceptance. Promotion requires the same frozen source plus six execution
records covering both Linux libc hosts per architecture, native macOS ARM64 and
macOS x64 through Rosetta. Every record must bind the actual binary SHA, real 0/7
status, default/ignored SIGCHLD, and ENOENT followed by a successful cleanup frame
and zero helper exit. The promotion command embeds that evidence in the manifest
and regenerates the bound MSBuild hash list. This marks the native asset set ready
for packaging; package, managed-runner and release acceptance remain separate gates.

Linux uses GCC 14.2.0 from Docker official images:

- x64: `gcc@sha256:82549aa8f90ada3236a8be70c74543132a76662ef33f0c3271ed802b81584a82`
- ARM64: `gcc@sha256:ab78e7a512cead0e90364bb21f91c299a913e8b19fc481bf700d97218524f46d`
- Official image source revision: `docker-library/gcc` commit
  `22c26a9e4edb4ce4f7676dfa100afc80fc2ccbea`, directory `14`.
- musl 1.2.5 source: <https://musl.libc.org/releases/musl-1.2.5.tar.gz>,
  SHA-256 `a9a118bbe84d8764da0ea0d28b3ab3fae8477fc7e4085d90102b8596fc7c75e4`.

musl is built from that archive with `--disable-shared`. The exact container
commands and compiler/linker versions are retained in the manifest. The final
link uses `-nostdlib -static` and explicitly supplies only musl's `crt1.o`,
`crti.o`, `libc.a`, and `crtn.o`. Retained link maps (temporary object paths and trailing whitespace normalized)
and an empty `nm -u` output
establish absence of compiler support libraries. The verifier independently
parses each ELF64 machine/type/program-header table and rejects an interpreter
or dynamic segment. The minimum Linux kernel has not been certified. Execution evidence states
the actual kernel and libc of each tested machine/container; those runs do not
claim support on untested historical kernels.

macOS uses CLT `26.6.0.0.1781586589`, Apple clang 21.0.0
(`clang-2100.1.1.101`), and macOS SDK 26.5 build `25F70`. The manifest also pins
the actual clang/linker and SDK settings hashes. x64 targets macOS 10.15; ARM64
targets macOS 11.0. The builder replaces `LC_UUID` with the first 16 bytes of SHA-256 over
`sourceSha256 + ":" + rid` before signing; recent dyld requires that command.
Explicit ad-hoc signing has no
signing timestamp and uses a constant identifier. The verifier parses the thin
Mach-O CPU, executable type, SDK/minimum OS and system dylibs. These deployment
minima are header/ABI constraints, not proof of a run on those historical OS
releases. On Apple Silicon, an x64 execution through Rosetta must be recorded as
such. No Apple system library or SDK is redistributed.

The musl notice retains the complete upstream COPYRIGHT (6204 bytes, SHA-256
`f9bc4423732350eb0b3f7ed7e91d530298476f8fec0c6c427a1c04ade22655af`) and
an appendix of individual AArch64 string-source copyright headers from the same
archive. Both the original MIT text and Arm notices are preserved. The verifier
pins the combined file bytes as well as the manifest license hash. GCC is a
build tool only; no libgcc/compiler support library is linked into these assets.

`NativeSupervisorAssetManifest.g.cs` and `SupervisorAssetHashes.props` are
committed generated files. Normal Core build/pack uses the .NET SDK's built-in
`GetFileHash` task against four binaries, C source, manifest, retained license,
and generated C# table. It does not invoke Python, a C compiler, Docker, or a shell.
The Python verifier is an additional maintainer/CI check and can optionally emit
`--write-csharp <path>` only after verification. Both Core target frameworks must
embed the complete four-RID set; RID selection occurs only when extracting the
one required executable at runtime.

Final-source native execution records are kept in
`artifacts/pr23-process-reaper/native-assets/`; the reviewed promotion embeds the
full six-environment proof and harness in `assets/manifest.json`. Source SHA-256
`4fab29e88329fa58530273a684c0bf3f95aca5f62ef9bf1eccb1c01f16986973` was
built twice per RID with byte-identical results, then executed in all six hosts:

| Asset | Actual execution environment | Result |
| --- | --- | --- |
| linux-x64 | WSL2 kernel 5.15.167.4, Ubuntu glibc 2.39 | 5/5 |
| linux-x64 | Same native x64 host, Alpine musl 1.2.6 container | 5/5 |
| linux-arm64 | OrbStack kernel 7.0.14, Debian glibc 2.36 | 5/5 |
| linux-arm64 | Same native ARM64 host, Alpine musl 1.2.6 container | 5/5 |
| osx-arm64 | macOS 26.6.2 build 25G83, ARM64 native | 5/5 |
| osx-x64 | Same macOS host, x64 executable through Rosetta | 5/5 |

Each host preserved real 0/7 exits under both default and inherited ignored
SIGCHLD. ENOENT preserved errno 2, then emitted CLEANED(0), and the supervisor
itself exited 0 in every case. Prior Darwin startup-cleanup failures were fixed
before this source freeze and these runs were repeated on the rebuilt assets.
Fourteen negative verifier probes rejected missing RIDs, changed source/binary
bytes, wrong native architectures, missing execution environments, failed helper
cleanup, wrong execution binary hashes, and altered real target exit status.

The two Linux libc hosts execute the same respective immutable binary. The
shipped libc inside each binary remains musl 1.2.5; musl 1.2.6 above identifies
the independent host-container runtime. Historical OS deployment minima,
managed-runner/PID-1 adversarial regressions, Windows Job behavior, and final
Ubuntu-built package consumers remain separate gates owned by the delivery
coordinator. These native-asset checks do not substitute for those gates.
