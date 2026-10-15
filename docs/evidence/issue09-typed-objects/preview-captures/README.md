# Native macOS Preview captures

These eight JPEGs are the unmodified bytes returned by the documented CUA `preview.getScreenshot` API, captured from the actual macOS Preview windows after confirming their native document URL and page count. Adjacent `.ax.txt` files preserve that exact native state. Each original generated PDF was reopened, its stated page rechecked and captured; the products were not regenerated. All five opened PDF windows were closed afterward.

| PDF | Native captures |
| --- | --- |
| before.pdf | [page 1](before-1.jpg), [page 2](before-2.jpg) |
| after.pdf | [page 1](after-1.jpg), [page 2](after-2.jpg) |
| original-v1.pdf | [page 1](original-v1-1.jpg), [page 2](original-v1-2.jpg) |
| binary-xml-split.pdf | [page 1](binary-xml-split-1.jpg) |
| split-template-liveness.pdf | [page 1](split-template-liveness-1.jpg) |

The [machine acceptance record](../preview-acceptance-2026-10-15.json) maps these captures to exact PDF identities and pages. The [member manifest](../member-manifest.json) records their lengths and hashes. Existing `sample/*-N.png` and `ci-fixture-repair/*-N.png` remain auxiliary Poppler renders. Window sizes differ naturally; no image resizing, cropping, composition, recompression or regeneration was applied. The file extension follows the actual JPEG format returned by CUA.
