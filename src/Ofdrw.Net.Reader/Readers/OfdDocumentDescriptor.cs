using System;
using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Reader.Readers;

/// <summary>A DocBody declaration in OFD.xml, in source order. It does not attest to document or signature validity.</summary>
public sealed class OfdDocumentDescriptor
{
    internal OfdDocumentDescriptor(int index, string documentEntryPath, string? title, IReadOnlyList<string> signatureListPaths)
    { Index = index; DocumentEntryPath = documentEntryPath; Title = title; SignatureListPaths = Array.AsReadOnly(signatureListPaths.ToArray()); }

    public int Index { get; }
    public string DocumentEntryPath { get; }
    public string? Title { get; }
    public IReadOnlyList<string> SignatureListPaths { get; }
}
