# Strongly typed standard object editing

`OfdReader.ReadAsync` exposes the selected document's additional standard objects through `OfdDocumentPackage.Options.Metadata`, `Permissions`, `ViewingPreferences`, `Versions`, and `Actions`. `OfdPackageLoadOptions.DocumentIndex` selects a DocBody by zero-based index. The default remains the first document. The [multi-DocBody edit boundary](multi-docbody.md) applies to every write.

## DocInfo

`OfdMetadata.DocUsage` is an optional `OfdDocUsage` value (`Normal`, `EBook`, `ENewsPaper`, `EMagzine`). `Cover` is a package-relative image location, `CreatorVersion` is text, and `CustomDatas` contains ordered `CustomData` entries with `Name` and text value. An omitted `DocUsage` retains the standard default. A newly assigned or changed `Cover` must resolve to an existing package entry; this library neither fetches nor creates the image.

```csharp
var reader = new OfdReader();
var document = await reader.ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
document.Options.Metadata.DocUsage = OfdDocUsage.EBook;
document.Options.Metadata.CustomDatas ??= new OfdCustomDatas();
document.Options.Metadata.CustomDatas.Add("case", "001");
await new OfdPackageWriter().WriteAsync(document, output);
```

`OfdCustomDatas.Items` and `OfdVersions.Items` are ordered views. Editing an item changes its parent; `Add`, `RemoveAt` and `Clear` manage modeled entries. An empty wrapper with no extension content is omitted at write time. Unknown attributes and children remain in their original positions. A malformed or ambiguous source node stays as preserved XML and cannot be replaced through a new typed instance without an explicit failure.

## Permissions, preferences and versions

`OfdPermissions` models the six optional Boolean declarations plus `Print` (`Printable`, `Copies`) and `ValidPeriod` (`StartDate`, `EndDate` XML dateTime strings). Nullable Boolean properties preserve omission separately from the standard's defaults. `OfdViewingPreferences` models page mode, layout, tab display, hide flags, and the exclusive `ZoomMode` or numeric `Zoom` choice. `OfdVersions` models the selected DocBody's ordered `Version` descriptors (`Id`, `Index`, optional `Current`, `BaseLocation`). Changed or new version locations must resolve to existing package entries. Version files themselves are preserved, not generated.

Permissions are **declarations only**. The Reader, Writer and editing APIs do not enforce them as access control. Viewer preferences likewise do not control a UI.

## Actions and source preservation

Existing local `OfdGoToAction` remains available. `OfdUriAction`, `OfdAttachmentAction` (`GotoA`), `OfdSoundAction`, and `OfdMovieAction` model the other standard action forms. URI, attachment and media actions are serialized only; the library does not request a URI, open an attachment or play media. Attachment and media targets must be retained local resources for new or edited typed actions. Unsupported, ambiguous or unprovable source actions remain `OfdRawAction` with their full XML.

Imported wrappers preserve unknown XML, namespaces and unmodeled siblings. Assigning null to a pristine imported typed wrapper preserves its source. Withdrawing a modified imported wrapper fails. A typed node from another document cannot be installed into the selected document. Page deletion or cross-document migration still fails when reference closure cannot be proven. This API provides targeted editing of the named objects, not whole-document XSD validation or archival certification.

`OfdDocumentSplitter.Split` carries edited custom metadata, permissions, viewer preferences, and version descriptors into the selected package with new document ownership. `OfdDocumentMerger.Merge` carries the first source's `DocUsage`, `Cover`, and `CreatorVersion`; a cover image must be an existing, distinct package entry so its bytes can be copied into the merged package.

The field shapes and order were checked against the public OFD [OFD.xsd](https://github.com/zc310/ofd/blob/main/internal/schema/xsd/OFD.xsd), [Document.xsd](https://github.com/zc310/ofd/blob/main/internal/schema/xsd/Document.xsd), and [Definitions.xsd](https://github.com/zc310/ofd/blob/main/internal/schema/xsd/Definitions.xsd), with the [upstream Java object model](https://github.com/ofdrw/ofdrw/tree/84e9eb260ba6fd351ecccf08a01a40305d4af635/ofdrw-core/src/main/java/org/ofdrw/core) as an interoperability cross-check.
