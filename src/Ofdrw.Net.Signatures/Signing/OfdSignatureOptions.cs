using System;
using System.Collections.Generic;
using Ofdrw.Net.Signatures.Crypto;

namespace Ofdrw.Net.Signatures.Signing;

public sealed class OfdSignatureOptions
{
    public string CheckMethod { get; set; } = OfdDigestAlgorithms.Sm3Oid;
    public bool ProtectSignatureList { get; set; }

    /// <summary>
    /// Optional veto for package entries. Receives an absolute package path such as
    /// /Doc_0/Document.xml. Required document content cannot be excluded.
    /// The default signs every available entry except the mutable signature list.
    /// </summary>
    public Func<string, bool>? ProtectFile { get; set; }

    /// <summary>OFD 2.0 ID of an earlier signature on which the new signature depends.</summary>
    public string? RelativeSignatureId { get; set; }

    /// <summary>OFD 2.0 signed extension properties, serialized under SignedInfo.</summary>
    public IList<OfdSignatureParameter> Parameters { get; } = new List<OfdSignatureParameter>();
}
