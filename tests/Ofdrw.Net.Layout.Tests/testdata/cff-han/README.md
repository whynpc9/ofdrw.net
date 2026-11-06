# Ofdrw CI Han CFF test face

`OfdrwCIHanCFF-Regular.otf` (87,168 bytes, SHA-256
`d3d866041ce45bc66aa51259c2dece59c9dc912140518f59e94b3fb061bebb3d`) is a
fontTools 4.44.0 subset of `Sans/SubsetOTF/SC/NotoSansSC-Regular.otf` from
`notofonts/noto-cjk` at commit `f8d157532fbfaeda587e826d4cd5b21a49186f7c`
(8,331,336 bytes, SHA-256
`faa6c9df652116dde789d351359f3d7e5d2285a2b2a1f04a2d7244df706d5ea9`).

It keeps the upstream CID-keyed CFF1 program (ROS, 18-entry FDArray, FDSelect,
global and per-FD local subroutines), GSUB/GPOS and an OS/2.fsType of zero, and
covers printable ASCII, about one hundred Han characters used by the layout
tests, U+3000 and eight CJK Extension B characters (U+20087, U+20089, U+200CC,
U+20164, U+20628, U+20676, U+20CD0, U+2139A). The family is renamed
`Ofdrw CI Han CFF` so the modified subset is never mistaken for the upstream
font. `LICENSE` is the unmodified upstream SIL OFL 1.1 text (SHA-256
`6a73f9541c2de74158c0e7cf6b0a58ef774f5a780bf191f2d7ec9cc53efe2bf2`).

`gid-*.path` are independent fontTools 4.44.0 `RecordingPen` outlines of the
original GIDs 72 (`g`), 257 (`中`), 263 (`体`) and 459 (U+20087), in design
units with three decimals: `M`, `L`, `C` (two controls and the end point) and
`Z`. The managed Type 2 interpreter is compared against them. Neither fontTools
nor Python is a product runtime dependency.
