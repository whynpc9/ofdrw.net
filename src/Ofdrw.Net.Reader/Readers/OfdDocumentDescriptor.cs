using System;
using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Reader.Readers;

/// <summary>A DocBody declaration in OFD.xml, in source order. It does not attest to document or signature validity.</summary>
public sealed class OfdDocumentDescriptor
{
    internal OfdDocumentDescriptor(int index, string documentEntryPath, string? title, IReadOnlyList<string> signatureListPaths)
    { Index = index; DocumentEntryPath = documentEntryPath; Title = title; SignatureListPaths = Array.AsReadOnly(signatureListPaths.ToArray()); }

    /// <summary>Zero-based position of the DocBody in OFD.xml.</summary>
    public int Index { get; }
    /// <summary>Normalized archive path to the document entry.</summary>
    public string DocumentEntryPath { get; }
    /// <summary>Optional DocInfo title.</summary>
    public string? Title { get; }
    /// <summary>Declared signature-list archive paths; these do not attest to signature validity.</summary>
    public IReadOnlyList<string> SignatureListPaths { get; }
}
