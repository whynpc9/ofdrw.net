# R10 self-authored generic-region fixture

`generate_owned_pattern.py` creates `owned-pattern.pbm` (64 x 56, PBM P4,
SHA-256 `24b65cb0dd3f64298301fe5ec33e5ff4600d598056351c28fe1f3b5e5917e1bc`).
The pattern is a formula written for this repository; it is not a copied test
image. `owned-generic-region.jb2` is its 173-byte standalone, single-page JBIG2
encoding (SHA-256 `16d262d9b7a3bd9856d42c8cf576a692bf96182854d3d7c0f5084ef9cad8a83f`).

The encoding was produced twice, byte-identically, by
[`agl/jbig2enc`](https://github.com/agl/jbig2enc) 0.32 at commit
`4553595aa094cad7d55c4efc9657ea31db23c742`. The bitstream was decoded
independently with [`jbig2dec`](https://github.com/ArtifexSoftware/jbig2dec)
0.20 and compared byte-for-byte with the PBM,
including polarity. The local reproduction commands and result hashes are in
`artifacts/roadmap-v12-20261021/acceptance/R10/4d5b7028ba51e4d69d803d3bba03fe380ec756db/owned-fixture-probe/`.
The encoder and independent decoder are development tools only; neither is a
runtime dependency nor bundled in NuGet packages. The separate ITU-derived
`sampledata_page1.jb2` was excluded because its redistribution grant is unknown.
