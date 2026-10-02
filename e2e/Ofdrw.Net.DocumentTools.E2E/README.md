# Document tool API and CLI examples

The program exercises all four public tools separately. It copies the source DOCX and changes only font declarations to the repository-pinned OFL Noto font, allowing the embedded-font evidence to be redistributed. The copied source and font/license are included in the evidence bundle. Native keeps viewer-local CJK fonts name-only, so the example explicitly embeds the same licensed measurement face through `OfdFontResource.Data` before tool operations; it does not change the converter default. It starts with the synthetic `generated-layout.docx` in explicit Native and default modes, adds a shared template, visible annotation, public attachment and a deliberately noncryptographic test signature/appearance, then exports each result to PDF/SVG and extracts OFD text.

```bash
python3 -m pip install fonttools==4.59.2
python3 scripts/install-ci-fonts.py --directory artifacts/document-tools-fonts
./scripts/run-document-tools-e2e.sh
```

Use the SDK selected by `global.json`; set a writable `DOTNET_CLI_HOME` and an existing `NUGET_PACKAGES` cache as appropriate. The script sets sandbox build flags. It requires Poppler and `rsvg-convert` for automatic rendering. CLI examples use one-based pages; the C# API uses zero-based positions. A synthetic red image is both an ordinary watermark and a signature-appearance fixture, without private data or production signing.

The Mix green TOP LAYER box covers the earlier blue UNDER LAYER box; first-page size and source ordering are intentional. Split orders source pages 2,1. Clean output retains every non-signature entry's bytes; the red signature square at (175,240) must disappear while body, template, annotation and attachment remain. The watermark red square at (175,215) is ordinary content and remains visible after merge and export.

The script also checks OFD/PDF text counts, selected page order, clean body byte equality and all file hashes. Automated rendering and PNG review do not complete macOS Preview acceptance. See [the persistent evidence record](../../docs/evidence/document-tools/README.md).

After extracting the evidence bundle, reproduction can use `OFDRW_DOCUMENT_TOOLS_FONTS=docs/evidence/document-tools/files/fonts` to reuse the exact licensed font without another download.

The annotation-clipped pair adds an oversized sheared image inside a 20x10 mm Appearance, nested PageBlock content, a 90-degree rotated ROTATE label and a 1.5x SCALE label. The clipped output and its Mix roundtrip must retain the exact parallelogram, glyph orientation/size and extract each label once.

`italic-fixed-anchor` uses original MIT rectangle glyphs with identical baseline anchors. The right column applies an unmarked shear to the whole glyph; its upper regular rectangle and lower italic rectangle must gain slope relative to the left identity column. `split-template-liveness` keeps an extensionless XML dependency resolved through a nested extension `BaseLoc` while selecting only page 2.

`resource-suffix-input` renames the real Noto font descriptor to `PublicResources.dat` and the ordinary watermark image descriptor to `ImageResources.bin`. `watermark-resource-suffix` adds RESOURCE SUFFIX while retaining the prior DRAFT label, red watermark image, body, annotation and template. Its two-page PDF and first-page SVG verify that declared resource XML needs no `.xml` suffix.
