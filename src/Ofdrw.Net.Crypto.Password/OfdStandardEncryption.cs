using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Crypto.Password.Internal;

namespace Ofdrw.Net.Crypto.Password;

/// <summary>Explicit, bounded Java-compatible single-layer SM4 envelope with multiple same-kind recipients.
/// Standard OFDEntries signs the file-name inventory; this package's signed digest extension additionally
/// authenticates each stored file byte sequence. Java-only inventories require explicit unverified-content opt-in.</summary>
public static class OfdStandardEncryption
{
    private static readonly XNamespace Ns = "http://www.ofdspec.org/2016";
    private static readonly XNamespace Ext = "urn:ofdrw-net:standard-encryption-content-digests:1";
    private const string EnvelopePath = "Encryptions.xml";
    private const string SeedPath = "decryptseed.dat";
    private const string MapPath = "entriesmap.dat";
    private const string ManifestPath = "OFDEntries.xml";
    private const string SignedValuePath = "signedvalue.dat";

    public static async Task<byte[]> EncryptAsync(Stream source, OfdStandardEncryptionOptions options,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(source); ArgumentNullException.ThrowIfNull(options);
        var settings = new OfdPasswordEnvelope.Settings(options.Limits);
        if (settings.Paths.Length != 0 || settings.Pages.Length != 0)
            throw new ArgumentException("Standard profile encrypts every file entry; selection is unsupported.", nameof(options));
        var recipients = options.Recipients.ToArray();
        if (recipients.Length == 0 || recipients.Length > 32 || recipients.Any(x => x is null))
            throw new ArgumentException("Provide 1 to 32 recipients.", nameof(options));
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var recipient in recipients)
            if (!names.Add(recipient.Name)) throw new ArgumentException("Recipient names must be unique.", nameof(options));
        bool certificate = recipients[0].CertificateDer is not null;
        if (recipients.Any(x => (x.CertificateDer is not null) != certificate))
            throw new ArgumentException("Mixed password and certificate recipients are outside this profile.", nameof(options));
        var signer = options.ManifestSigner ?? throw new ArgumentException("A pinned manifest signer is required.", nameof(options));
        using var owned = new OwnedBuffers();
        var archive = await OfdPasswordEnvelope.LoadRawValidatedAsync(source, settings, cancellationToken).ConfigureAwait(false);
        var entries = archive.EntryNames.ToDictionary(path => path, archive.GetBytes, StringComparer.OrdinalIgnoreCase);
        OfdPasswordEnvelope.AdoptAndValidateEntries(entries, owned);
        if (entries.Keys.Any(Reserved)) throw new InvalidDataException("Source already has encryption or integrity metadata.");
        _ = OfdPasswordEnvelope.ValidateDocument(entries, settings);
        var original = entries.Keys.OrderBy(x => x, StringComparer.Ordinal).ToArray();
        var fek = RandomNumberGenerator.GetBytes(16); var iv = RandomNumberGenerator.GetBytes(16);
        owned.Add(fek); owned.Add(iv);
        var map = new XElement(Ns + "EncryptEntries", new XAttribute("ID", "1"));
        for (int i = 0; i < original.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cipherPath = $"Cipher/entry-{i + 1:D5}.dat";
            var cipher = PasswordPrimitives.Cbc(true, entries[original[i]], fek, iv, cancellationToken);
            owned.Add(cipher); entries.Remove(original[i]); entries.Add(cipherPath, cipher);
            map.Add(new XElement(Ns + "EncryptEntry", new XAttribute("Path", "/" + original[i]),
                new XAttribute("EPath", "/" + cipherPath)));
        }
        var mapBytes = OfdPasswordEnvelope.XmlBytes(map, settings.Metadata, owned);
        var mapCipher = PasswordPrimitives.Cbc(true, mapBytes, fek, iv, cancellationToken);
        owned.Add(mapCipher); entries.Add(MapPath, mapCipher);
        var seed = new XElement(Ns + "DecyptSeed", new XAttribute("ID", "1"),
            new XAttribute("EncryptCaseId", certificate ? "1.1.2" : "1.1.1"));
        foreach (var recipient in recipients)
        {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] wrapped;
            if (certificate) wrapped = StandardCryptoPrimitives.EncryptSm2(fek, recipient.CertificateDer!);
            else
            {
                var kek = StandardCryptoPrimitives.PasswordKey(recipient.Password!); owned.Add(kek);
                wrapped = PasswordPrimitives.Cbc(true, fek, kek, iv, cancellationToken);
            }
            owned.Add(wrapped);
            var user = new XElement(Ns + "UserInfo", new XAttribute("UserName", recipient.Name),
                new XAttribute("UserType", "User"));
            if (certificate) user.Add(new XElement(Ns + "UserCert", Convert.ToBase64String(recipient.CertificateDer!)));
            user.Add(new XElement(Ns + "EncryptedWK", Convert.ToBase64String(wrapped)),
                new XElement(Ns + "IVValue", Convert.ToBase64String(iv)));
            seed.Add(user);
        }
        seed.Add(new XElement(Ns + "ExtendParams"));
        entries.Add(SeedPath, OfdPasswordEnvelope.XmlBytes(seed, settings.Metadata, owned));
        var envelope = new XElement(Ns + "Encryptions", new XElement(Ns + "EncryptInfo", new XAttribute("ID", "1"),
            new XElement(Ns + "Provider", new XAttribute("Name", "Ofdrw.Net.Crypto.Password"),
                new XAttribute("Company", "Ofdrw.Net"), new XAttribute("Version", "1")),
            new XElement(Ns + "EncryptScope", "All"), new XElement(Ns + "DecryptSeedLoc", "/" + SeedPath),
            new XElement(Ns + "EntriesMapLoc", "/" + MapPath)));
        entries.Add(EnvelopePath, OfdPasswordEnvelope.XmlBytes(envelope, settings.Metadata, owned));
        var manifest = CreateManifest(entries, settings, owned);
        entries.Add(ManifestPath, manifest);
        var signedValue = await StandardManifestSignature.SignAsync(manifest, signer, cancellationToken).ConfigureAwait(false);
        owned.Add(signedValue); entries.Add(SignedValuePath, signedValue);
        OfdPasswordEnvelope.CheckExpanded(entries, settings);
        return await OfdPasswordEnvelope.Pack(entries, settings, cancellationToken).ConfigureAwait(false);
    }

    public static async Task<OfdStandardDecryptionResult> DecryptAsync(Stream source, OfdStandardCredential credential,
        OfdStandardDecryptionOptions options, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(source); ArgumentNullException.ThrowIfNull(credential);
        ArgumentNullException.ThrowIfNull(options);
        var settings = new OfdPasswordEnvelope.Settings(options.Limits);
        if (settings.Paths.Length != 0 || settings.Pages.Length != 0)
            throw new ArgumentException("Standard profile does not use v2 selection settings.", nameof(options));
        using var owned = new OwnedBuffers();
        var archive = await OfdPasswordEnvelope.LoadRawValidatedAsync(source, settings, cancellationToken).ConfigureAwait(false);
        var entries = archive.EntryNames.ToDictionary(path => path, archive.GetBytes, StringComparer.OrdinalIgnoreCase);
        OfdPasswordEnvelope.AdoptAndValidateEntries(entries, owned);
        if (!entries.ContainsKey(EnvelopePath) || !entries.ContainsKey(SeedPath) || !entries.ContainsKey(MapPath))
            throw new InvalidDataException("Required standard encryption metadata is missing.");
        var contentVerified = VerifyManifest(entries, options, settings, cancellationToken);
        var envelope = OfdPasswordEnvelope.Parse(OfdPasswordEnvelope.Get(entries, EnvelopePath), settings);
        Need(envelope.Name == Ns + "Encryptions" && envelope.Attributes().All(a => a.IsNamespaceDeclaration) &&
             envelope.Elements().Count() == 1, "Unsupported Encryptions.xml structure.");
        var info = One(envelope, Ns + "EncryptInfo");
        Need((string?)info.Attribute("ID") == "1" && AttrsOnly(info, "ID") &&
            info.Elements().All(x => x.Name == Ns + "Provider" || x.Name == Ns + "EncryptScope" ||
                x.Name == Ns + "DecryptSeedLoc" || x.Name == Ns + "EntriesMapLoc" || x.Name == Ns + "EncryptDate"),
            "Unsupported encryption information.");
        var provider = One(info, Ns + "Provider");
        Need(AttrsOnly(provider, "Name", "Company", "Version") &&
            !string.IsNullOrWhiteSpace((string?)provider.Attribute("Name")) &&
            !string.IsNullOrWhiteSpace((string?)provider.Attribute("Company")) &&
            !string.IsNullOrWhiteSpace((string?)provider.Attribute("Version")) && !provider.HasElements,
            "Unsupported encryption provider.");
        Need(Leaf(info, "EncryptScope") == "All" && Absolute(Leaf(info, "DecryptSeedLoc")) == SeedPath &&
            Absolute(Leaf(info, "EntriesMapLoc")) == MapPath, "Unsupported encryption scope or metadata paths.");
        var seed = OfdPasswordEnvelope.Parse(OfdPasswordEnvelope.Get(entries, SeedPath), settings);
        Need(seed.Name == Ns + "DecyptSeed" && (string?)seed.Attribute("ID") == "1" &&
            AttrsOnly(seed, "ID", "EncryptCaseId") &&
            seed.Elements().All(x => x.Name == Ns + "UserInfo" || x.Name == Ns + "ExtendParams"),
            "Unsupported key descriptor.");
        var extend = seed.Elements(Ns + "ExtendParams").ToArray();
        Need(extend.Length <= 1 && (extend.Length == 0 || (!extend[0].HasAttributes && !extend[0].HasElements &&
            string.IsNullOrWhiteSpace(extend[0].Value))), "Unsupported key extension parameters.");
        var caseId = (string?)seed.Attribute("EncryptCaseId");
        bool certificate = credential.CertificateDer is not null;
        Need(caseId is null || caseId == (certificate ? "1.1.2" : "1.1.1"), "Recipient scheme does not match key descriptor.");
        var users = seed.Elements(Ns + "UserInfo").ToArray();
        Need(users.Length is >= 1 and <= 32 && users.Select(x => (string?)x.Attribute("UserName"))
            .All(x => !string.IsNullOrWhiteSpace(x)) &&
            users.All(x => AttrsOnly(x, "UserName", "UserType") &&
                x.Elements().All(c => c.Name == Ns + "UserCert" || c.Name == Ns + "EncryptedWK" || c.Name == Ns + "IVValue") &&
                (string?)x.Attribute("UserType") is null or "User" or "Owner") &&
            users.Select(x => (string?)x.Attribute("UserName")).Distinct(StringComparer.OrdinalIgnoreCase).Count() == users.Length,
            "Missing or ambiguous UserInfo.");
        byte[]? commonIv = null;
        foreach (var candidate in users)
        {
            Need(candidate.Elements(Ns + "EncryptedWK").Count() == 1 &&
                candidate.Elements(Ns + "IVValue").Count() == 1 &&
                candidate.Elements(Ns + "UserCert").Count() == (certificate ? 1 : 0),
                "Mixed or incomplete recipient schemes.");
            var candidateIv = Decode(Leaf(candidate, "IVValue"), 16, 16);
            try
            {
                Need(commonIv is null || CryptographicOperations.FixedTimeEquals(commonIv, candidateIv),
                    "Recipients do not share one package IV.");
                commonIv ??= (byte[])candidateIv.Clone();
            }
            finally { CryptographicOperations.ZeroMemory(candidateIv); }
        }
        if (commonIv is not null) owned.Add(commonIv);
        var user = users.SingleOrDefault(x => (string?)x.Attribute("UserName") == credential.Name)
            ?? throw new InvalidDataException("Recipient is absent.");
        Need(user.Elements().All(x => x.Name == Ns + "UserCert" || x.Name == Ns + "EncryptedWK" || x.Name == Ns + "IVValue"),
            "Unsupported UserInfo.");
        var iv = Decode(Leaf(user, "IVValue"), 16, 16); owned.Add(iv);
        var wrapped = Decode(Leaf(user, "EncryptedWK"), certificate ? 64 : 32, certificate ? 256 : 32);
        owned.Add(wrapped);
        byte[] fek;
        if (certificate)
        {
            var embedded = Decode(Leaf(user, "UserCert"), 64, 64 * 1024); owned.Add(embedded);
            Need(CryptographicOperations.FixedTimeEquals(embedded, credential.CertificateDer!),
                "Recipient certificate does not match the pinned certificate.");
            fek = StandardCryptoPrimitives.DecryptSm2(wrapped, credential.PrivateKeyPkcs8Der!);
        }
        else
        {
            Need(!user.Elements(Ns + "UserCert").Any(), "Unexpected certificate in password recipient.");
            var kek = StandardCryptoPrimitives.PasswordKey(credential.Password!); owned.Add(kek);
            fek = PasswordPrimitives.Cbc(false, wrapped, kek, iv, cancellationToken);
            Need(fek.Length == 16, "Invalid recovered file key.");
        }
        owned.Add(fek);
        var mapBytes = PasswordPrimitives.Cbc(false, OfdPasswordEnvelope.Get(entries, MapPath), fek, iv, cancellationToken);
        owned.Add(mapBytes);
        var map = OfdPasswordEnvelope.Parse(mapBytes, settings);
        Need(map.Name == Ns + "EncryptEntries" && (string?)map.Attribute("ID") == "1" && AttrsOnly(map, "ID") &&
            map.Elements().All(x => x.Name == Ns + "EncryptEntry"), "Unsupported encrypted map.");
        var mappings = map.Elements(Ns + "EncryptEntry").ToArray();
        Need(mappings.Length > 0 && mappings.Length <= settings.Load.MaxEntryCount, "Empty or oversized encrypted map.");
        var clear = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var cipher = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var mapping in mappings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Need(AttrsOnly(mapping, "Path", "EPath") && mapping.Attributes().Count() == 2 && !mapping.HasElements,
                "Unsupported map entry.");
            var plainPath = Absolute((string?)mapping.Attribute("Path"));
            var encrypted = Absolute((string?)mapping.Attribute("EPath"));
            Need(!Reserved(plainPath) && !Metadata(encrypted) && clear.Add(plainPath) && cipher.Add(encrypted) &&
                !entries.ContainsKey(plainPath), "Colliding or duplicate encrypted map entry.");
        }
        var expected = new HashSet<string>(cipher, StringComparer.OrdinalIgnoreCase)
            { EnvelopePath, SeedPath, MapPath, ManifestPath, SignedValuePath };
        Need(expected.SetEquals(entries.Keys), "Encrypted package has missing or unexpected file entries.");
        foreach (var mapping in mappings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var plainPath = Absolute((string?)mapping.Attribute("Path"));
            var encrypted = Absolute((string?)mapping.Attribute("EPath"));
            var plain = PasswordPrimitives.Cbc(false, OfdPasswordEnvelope.Get(entries, encrypted), fek, iv, cancellationToken);
            owned.Add(plain); entries.Remove(encrypted); entries.Add(plainPath, plain);
        }
        entries.Remove(EnvelopePath); entries.Remove(SeedPath); entries.Remove(MapPath);
        entries.Remove(ManifestPath); entries.Remove(SignedValuePath);
        _ = OfdPasswordEnvelope.ValidateDocument(entries, settings);
        OfdPasswordEnvelope.CheckExpanded(entries, settings);
        var restored = await OfdPasswordEnvelope.Pack(entries, settings, cancellationToken).ConfigureAwait(false);
        return new OfdStandardDecryptionResult(restored, true, true, contentVerified);
    }

    /// <summary>Commit an encrypted package only after all processing succeeds.</summary>
    public static async Task EncryptFileAsync(string sourcePath, string destinationPath,
        OfdStandardEncryptionOptions options, CancellationToken cancellationToken = default)
    {
        await FileTransformAsync(sourcePath, destinationPath,
            input => EncryptAsync(input, options, cancellationToken), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Validate the complete encrypted package before publishing plaintext by same-directory rename.</summary>
    public static async Task<OfdStandardFileDecryptionStatus> DecryptFileAsync(string sourcePath, string destinationPath,
        OfdStandardCredential credential, OfdStandardDecryptionOptions options, CancellationToken cancellationToken = default)
    {
        OfdStandardFileDecryptionStatus? status = null;
        await FileTransformAsync(sourcePath, destinationPath, async input =>
        {
            var result = await DecryptAsync(input, credential, options, cancellationToken).ConfigureAwait(false);
            status = new(result.ManifestSignatureVerified, result.ExactEntrySetVerified, result.FileContentsVerified);
            return result.OfdBytes;
        }, cancellationToken).ConfigureAwait(false);
        return status!;
    }

    private static async Task FileTransformAsync(string sourcePath, string destinationPath,
        Func<Stream, Task<byte[]>> transform, CancellationToken token)
    {
        var source = Path.GetFullPath(sourcePath); var destination = Path.GetFullPath(destinationPath);
        if (source.Equals(destination, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Input and output files must differ.");
        byte[]? result = null; string? temporary = null;
        try
        {
            await using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
                result = await transform(input).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            temporary = Path.Combine(Path.GetDirectoryName(destination)!, ".ofd-standard-" + Guid.NewGuid().ToString("N") + ".tmp");
            var options = new FileStreamOptions { Mode = FileMode.CreateNew, Access = FileAccess.Write, Share = FileShare.None,
                Options = FileOptions.Asynchronous };
            if (!OperatingSystem.IsWindows()) options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            await using (var output = new FileStream(temporary, options))
            { await output.WriteAsync(result, token).ConfigureAwait(false); await output.FlushAsync(token).ConfigureAwait(false); }
            token.ThrowIfCancellationRequested();
            File.Move(temporary, destination, overwrite: true); temporary = null;
        }
        finally
        {
            if (result is not null) CryptographicOperations.ZeroMemory(result);
            if (temporary is not null && File.Exists(temporary)) File.Delete(temporary);
        }
    }

    private static byte[] CreateManifest(Dictionary<string, byte[]> entries,
        OfdPasswordEnvelope.Settings settings, OwnedBuffers owned)
    {
        var list = new XElement(Ns + "FileList");
        var hashes = new XElement(Ext + "ContentDigests", new XAttribute("Algorithm", "SM3"));
        int id = 1;
        foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.Ordinal))
        {
            list.Add(new XElement(Ns + "File", new XAttribute("ID", id++.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("FileLoc", "/" + entry.Key)));
            hashes.Add(new XElement(Ext + "Entry", new XAttribute("Path", "/" + entry.Key),
                new XAttribute("Digest", Convert.ToBase64String(PasswordPrimitives.Sm3(entry.Value)))));
        }
        var root = new XElement(Ns + "DocEntries", new XAttribute("ID", "1"),
            new XElement(Ns + "SignedValueLoc", "/" + SignedValuePath), list, hashes);
        return OfdPasswordEnvelope.XmlBytes(root, settings.Metadata, owned);
    }

    private static bool VerifyManifest(Dictionary<string, byte[]> entries, OfdStandardDecryptionOptions options,
        OfdPasswordEnvelope.Settings settings, CancellationToken token)
    {
        var expectedCert = options.ExpectedManifestSignerCertificateDer
            ?? throw new ArgumentException("Pin the expected manifest signer certificate.", nameof(options));
        var manifestBytes = OfdPasswordEnvelope.Get(entries, ManifestPath);
        var manifest = OfdPasswordEnvelope.Parse(manifestBytes, settings);
        Need(manifest.Name == Ns + "DocEntries" && !string.IsNullOrWhiteSpace((string?)manifest.Attribute("ID")) &&
            AttrsOnly(manifest, "ID", "CreatorName", "Version", "CreationDate") &&
            manifest.Elements().All(x => x.Name == Ns + "SignedValueLoc" || x.Name == Ns + "FileList" ||
                x.Name == Ext + "ContentDigests"), "Unsupported OFDEntries.xml.");
        var signedPath = Absolute(Leaf(manifest, "SignedValueLoc"));
        Need(signedPath == SignedValuePath, "Unsupported signed value location.");
        Need(StandardManifestSignature.Verify(manifestBytes, OfdPasswordEnvelope.Get(entries, SignedValuePath), expectedCert),
            "OFDEntries signature or signer is invalid.");
        var list = One(manifest, Ns + "FileList");
        Need(!list.HasAttributes && list.Elements().All(x => x.Name == Ns + "File"), "Unsupported file inventory.");
        var listed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var file in list.Elements())
        {
            token.ThrowIfCancellationRequested();
            var id = (string?)file.Attribute("ID");
            Need(id is not null && ids.Add(id) && AttrsOnly(file, "ID", "FileLoc") &&
                file.Attributes().Count() == 2 && !file.HasElements &&
                listed.Add(Absolute((string?)file.Attribute("FileLoc"))), "Duplicate or malformed file inventory.");
        }
        var actual = new HashSet<string>(entries.Keys, StringComparer.OrdinalIgnoreCase);
        actual.Remove(ManifestPath); actual.Remove(SignedValuePath);
        Need(listed.SetEquals(actual), "Manifest has missing or extra package entries.");
        var extensions = manifest.Elements(Ext + "ContentDigests").ToArray();
        Need(extensions.Length <= 1, "Duplicate content digest extension.");
        if (extensions.Length == 0)
        {
            Need(options.AllowUnverifiedContents, "Java file-name inventory does not authenticate file contents.");
            return false;
        }
        var digestList = extensions[0];
        Need((string?)digestList.Attribute("Algorithm") == "SM3" && AttrsOnly(digestList, "Algorithm") &&
            digestList.Elements().All(x => x.Name == Ext + "Entry"), "Unsupported content digest extension.");
        var hashed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in digestList.Elements())
        {
            token.ThrowIfCancellationRequested();
            var path = Absolute((string?)item.Attribute("Path"));
            Need(hashed.Add(path) && listed.Contains(path) && AttrsOnly(item, "Path", "Digest") &&
                item.Attributes().Count() == 2 && !item.HasElements,
                "Duplicate or malformed content digest.");
            var digest = Decode((string?)item.Attribute("Digest"), 32, 32);
            var computed = PasswordPrimitives.Sm3(entries[path]);
            try { Need(CryptographicOperations.FixedTimeEquals(digest, computed),
                "Stored file contents do not match signed digests."); }
            finally { CryptographicOperations.ZeroMemory(digest); CryptographicOperations.ZeroMemory(computed); }
        }
        Need(hashed.SetEquals(listed), "Signed content digest inventory is incomplete.");
        return true;
    }

    private static bool Metadata(string path) => path.Equals(EnvelopePath, StringComparison.OrdinalIgnoreCase) ||
        path.Equals(SeedPath, StringComparison.OrdinalIgnoreCase) || path.Equals(MapPath, StringComparison.OrdinalIgnoreCase) ||
        path.Equals(ManifestPath, StringComparison.OrdinalIgnoreCase) || path.Equals(SignedValuePath, StringComparison.OrdinalIgnoreCase);
    private static bool Reserved(string path) => Metadata(path) ||
        path.Equals("Cipher", StringComparison.OrdinalIgnoreCase) || path.StartsWith("Cipher/", StringComparison.OrdinalIgnoreCase) ||
        path.Equals("PasswordCrypto", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("PasswordCrypto/", StringComparison.OrdinalIgnoreCase);
    private static string Absolute(string? path)
    {
        Need(path is not null && path.StartsWith('/') && !path.StartsWith("//", StringComparison.Ordinal),
            "Expected absolute package path.");
        return OfdPasswordEnvelope.Canonical(path![1..]);
    }
    private static XElement One(XElement parent, XName name)
    {
        var all = parent.Elements(name).ToArray(); Need(all.Length == 1, "Expected one " + name.LocalName + "."); return all[0];
    }
    private static string Leaf(XElement parent, string name)
    {
        var node = One(parent, Ns + name); Need(!node.HasElements && !node.HasAttributes, "Malformed " + name + ".");
        return node.Value;
    }
    private static bool AttrsOnly(XElement element, params string[] allowed) =>
        element.Attributes().Where(x => !x.IsNamespaceDeclaration).All(x => x.Name.NamespaceName.Length == 0 &&
            allowed.Contains(x.Name.LocalName, StringComparer.Ordinal));
    private static byte[] Decode(string? encoded, int minimum, int maximum)
    {
        Need(encoded is not null && encoded.Length <= ((maximum + 2) / 3 * 4) + 4, "Invalid base64 field.");
        try { var value = Convert.FromBase64String(encoded!); Need(value.Length >= minimum && value.Length <= maximum,
            "Invalid binary field length."); return value; }
        catch (FormatException ex) { throw new InvalidDataException("Invalid base64 field.", ex); }
    }
    private static void Need(bool condition, string message)
    { if (!condition) throw new InvalidDataException(message); }
}
