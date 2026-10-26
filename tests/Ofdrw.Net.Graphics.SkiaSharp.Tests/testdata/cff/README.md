# Source Code Pro CFF1 test face

`SourceCodePro-Regular.otf` and `LICENSE.md` are the unmodified files from
Adobe's `source-code-pro` `release` branch at commit
`803b7e23ec97ae58b6232ea76519a76d428ba268`:

- `OTF/SourceCodePro-Regular.otf`: 131,128 bytes, SHA-256
  `9f9664e2edf6f045c11e774f9bd0be6993971f2544a39061a5ce478b96b051f8`
- `LICENSE.md`: 4,566 bytes, SHA-256
  `7c940e28a5388e9bba866cf0e408edda45fe0899ba98665b8f6ab31dc5e4b8ff`

The adjacent upstream license is SIL OFL 1.1 and names `Source` as a Reserved
Font Name. The face remains unmodified; its `OS/2.fsType` is zero. The OTF is
tracked through Git LFS. Development-only fontTools 4.60.1 generated the
independent CFF charstring commands in `gid-*.path` for original GIDs 2, 34,
298, and 998. Each line is a `RecordingPen` move, line, cubic, or close in
design units, formatted to three decimals; the focused tests pin each full
file's SHA-256. Neither fontTools nor Python is a product runtime dependency.
