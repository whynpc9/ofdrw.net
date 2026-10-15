using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Crypto.Password.Internal;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Crypto.Password;

/// <summary>Opt-in, single-layer password envelopes for this package's own profile. No vendor interoperability,
/// cryptographic authenticity, certification, or long-term archival suitability is implied.</summary>
public static class OfdPasswordEnvelope
{
    private const string EnvelopePath = "Encryptions.xml";
    private const string SeedPath = "PasswordCrypto/decryptseed.dat";
    private const string MapPath = "PasswordCrypto/entriesmap.dat";
    private const string ProviderName = "Ofdrw.Net.Crypto.Password";
    private static readonly XNamespace Ns = "http://www.ofdspec.org";
    private static readonly XNamespace Profile = "urn:ofdrw-net:password-profile:1";

    /// <summary>Encrypts an explicitly selected set of entries and returns a complete ZIP. Leaves input open.
    /// On failure no result is returned. The caller owns and should clear returned bytes when no longer needed.</summary>
    public static Task<byte[]> EncryptAsync(Stream source, string password, OfdPasswordOptions options,
        CancellationToken cancellationToken = default) => TransformAsync(source, password, options, true, cancellationToken);

    /// <summary>Decrypts only this package's versioned profile after verifying its recovery checksums.
    /// Wrong password and damaged ciphertext share InvalidDataException. Leaves input open; returns no partial plaintext.</summary>
    public static Task<byte[]> DecryptAsync(Stream source, string password, OfdPasswordOptions? options = null,
        CancellationToken cancellationToken = default) => TransformAsync(source, password, options ?? new(), false, cancellationToken);

    /// <summary>Encrypts to a distinct file through a same-directory atomic rename. Existing output survives failures before commit.</summary>
    public static Task EncryptFileAsync(string sourcePath, string destinationPath, string password, OfdPasswordOptions options,
        CancellationToken cancellationToken = default) => TransformFileAsync(sourcePath, destinationPath, password, options, true, cancellationToken);

    /// <summary>Validates all plaintext before writing a restricted temporary file and committing it with an atomic rename.
    /// Cancellation and ordinary exceptions delete the temporary file. Process termination and filesystem failure are outside this guarantee.</summary>
    public static Task DecryptFileAsync(string sourcePath, string destinationPath, string password, OfdPasswordOptions? options = null,
        CancellationToken cancellationToken = default) => TransformFileAsync(sourcePath, destinationPath, password, options ?? new(), false, cancellationToken);

    private static async Task<byte[]> TransformAsync(Stream source, string password, OfdPasswordOptions options, bool encrypt, CancellationToken token)
    {
        ArgumentNullException.ThrowIfNull(source);
        var settings = new Settings(options);
        token.ThrowIfCancellationRequested();
        var kek = PasswordPrimitives.PasswordKey(password);
        var owned = new OwnedBuffers();
        Dictionary<string, byte[]>? entries = null;
        try
        {
            var archive = await LoadRawValidatedAsync(source, settings, token).ConfigureAwait(false);
            entries = archive.EntryNames.ToDictionary(path => path, archive.GetBytes, StringComparer.OrdinalIgnoreCase);
            AdoptAndValidateEntries(entries, owned);
            if (encrypt) Encrypt(entries, kek, settings, owned, token);
            else Decrypt(entries, kek, settings, owned, token);
            token.ThrowIfCancellationRequested();
            return await Pack(entries, settings, token).ConfigureAwait(false);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(kek);
            owned.Dispose();
        }
    }

    private static async Task<OfdPackageArchive> LoadRawValidatedAsync(Stream source, Settings settings, CancellationToken token)
    {
        // Freeze precisely the bytes read from the caller's current position so
        // raw-name validation and the loader see one identical archive.
        using var snapshot = new BoundedMemoryStream(settings.Load.MaxInputBytes);
        if (source.CanSeek && source.Length - source.Position > settings.Load.MaxInputBytes)
            throw new InvalidDataException("OFD input exceeds the configured compressed input size limit.");
        var buffer = new byte[81920];
        try
        {
            long total = 0;
            while (true)
            {
                token.ThrowIfCancellationRequested();
                var remaining = settings.Load.MaxInputBytes - total;
                var count = remaining >= buffer.Length ? buffer.Length : (int)remaining + 1;
                var read = await source.ReadAsync(buffer, 0, count, token).ConfigureAwait(false);
                if (read == 0) break;
                Require(read <= remaining, "OFD input exceeds the configured compressed input size limit.");
                await snapshot.WriteAsync(buffer.AsMemory(0, read), token).ConfigureAwait(false);
                total += read;
            }
        }
        finally { CryptographicOperations.ZeroMemory(buffer); }
        snapshot.Position = 0;
        using (var zip = new ZipArchive(snapshot, ZipArchiveMode.Read, leaveOpen: true))
        {
            Require(zip.Entries.Count <= settings.Load.MaxEntryCount, "Raw ZIP entry count exceeds budget.");
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in zip.Entries)
            {
                token.ThrowIfCancellationRequested();
                var raw = entry.FullName;
                Require(names.Add(raw), "Raw ZIP contains duplicate entry names.");
                if (raw.EndsWith("/", StringComparison.Ordinal))
                {
                    Require(entry.Length == 0, "Directory entries must not contain payloads.");
                    Canonical(raw.Substring(0, raw.Length - 1));
                }
                else
                {
                    Canonical(raw);
                    var basename = raw.Substring(raw.LastIndexOf('/') + 1);
                    Require(!string.IsNullOrWhiteSpace(basename), "File entries require a nonblank basename.");
                }
            }
        }
        snapshot.Position = 0;
        return await new OfdPackageLoader().LoadAsync(snapshot, settings.Load, token).ConfigureAwait(false);
    }

    internal static void AdoptAndValidateEntries(Dictionary<string, byte[]> entries, OwnedBuffers owned)
    {
        // Register the entire collection before any path can reject it. A later
        // invalid name must not strand earlier or subsequent plaintext arrays.
        owned.AddRange(entries.Values);
        foreach (var path in entries.Keys) Canonical(path);
    }

    private static void Encrypt(Dictionary<string, byte[]> entries, byte[] kek, Settings settings, OwnedBuffers owned, CancellationToken token)
    {
        if (entries.Keys.Any(Reserved)) throw new InvalidDataException("Input already contains encryption metadata or a reserved PasswordCrypto path.");
        var root = ValidateDocument(entries, settings);
        var selected = SelectEntries(entries, root, settings);
        if (selected.Count == 0) throw new ArgumentException("Select at least one entry or page.");
        var fek = RandomNumberGenerator.GetBytes(16);
        var iv = RandomNumberGenerator.GetBytes(16);
        owned.Add(fek); owned.Add(iv);
        var map = new XElement(Ns + "EncryptEntries", new XAttribute("ID", "1"), new XAttribute(Profile + "Version", "1"));
        var inventory = new XElement(Profile + "Inventory");
        foreach (var entry in entries.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            token.ThrowIfCancellationRequested();
            inventory.Add(new XElement(Profile + "Entry", new XAttribute("Path", "/" + entry.Key),
                new XAttribute("Length", entry.Value.Length), new XAttribute("SM3", Convert.ToBase64String(PasswordPrimitives.Sm3(entry.Value)))));
        }
        int index = 0;
        foreach (var path in selected.OrderBy(path => path, StringComparer.Ordinal))
        {
            token.ThrowIfCancellationRequested();
            var cipherPath = $"PasswordCrypto/entry-{++index:D5}.dat";
            var cipher = PasswordPrimitives.Cbc(true, entries[path], fek, iv, token);
            owned.Add(cipher);
            entries.Remove(path); entries.Add(cipherPath, cipher);
            map.Add(new XElement(Ns + "EncryptEntry", new XAttribute("Path", "/" + path), new XAttribute("EPath", "/" + cipherPath)));
        }
        map.Add(inventory);
        var mapBytes = XmlBytes(map, settings.Metadata, owned);
        var encryptedMap = PasswordPrimitives.Cbc(true, mapBytes, fek, iv, token); owned.Add(encryptedMap);
        entries.Add(MapPath, encryptedMap);
        var wrapped = PasswordPrimitives.Cbc(true, fek, kek, iv, token); owned.Add(wrapped);
        var seed = new XElement(Ns + "DecyptSeed", new XAttribute("ID", "1"), new XAttribute("EncryptCaseId", "1.1.1"),
            new XElement(Ns + "UserInfo", new XAttribute("UserName", settings.UserName), new XAttribute("UserType", "User"),
                new XElement(Ns + "EncryptedWK", Convert.ToBase64String(wrapped)),
                new XElement(Ns + "IVValue", Convert.ToBase64String(iv))), new XElement(Ns + "ExtendParams"));
        var seedBytes = XmlBytes(seed, settings.Metadata, owned); entries.Add(SeedPath, seedBytes);
        var envelope = new XElement(Ns + "Encryptions", new XAttribute(Profile + "Version", "1"),
            new XElement(Ns + "EncryptInfo", new XAttribute("ID", "1"),
                new XElement(Ns + "Provider", new XAttribute("Name", ProviderName), new XAttribute("Company", "Ofdrw.Net"), new XAttribute("Version", "1")),
                new XElement(Ns + "EncryptScope", "Partial"),
                new XElement(Ns + "DecryptSeedLoc", "/" + SeedPath), new XElement(Ns + "EntriesMapLoc", "/" + MapPath)));
        var envelopeBytes = XmlBytes(envelope, settings.Metadata, owned); entries.Add(EnvelopePath, envelopeBytes);
        CheckExpanded(entries, settings);
    }

    private static void Decrypt(Dictionary<string, byte[]> entries, byte[] kek, Settings settings, OwnedBuffers owned, CancellationToken token)
    {
        var envelope = Parse(Get(entries, EnvelopePath), settings);
        Require(envelope.Name == Ns + "Encryptions" && (string?)envelope.Attribute(Profile + "Version") == "1", "Unsupported password profile.");
        Shape(envelope, [Profile + "Version"], [Ns + "EncryptInfo"]);
        var info = Single(envelope, Ns + "EncryptInfo");
        Shape(info, ["ID"], [Ns + "Provider", Ns + "EncryptScope", Ns + "DecryptSeedLoc", Ns + "EntriesMapLoc"]);
        Require((string?)info.Attribute("ID") == "1", "Unsupported encryption ID.");
        var provider = Single(info, Ns + "Provider");
        Shape(provider, ["Name", "Company", "Version"], []);
        Require((string?)provider.Attribute("Name") == ProviderName && (string?)provider.Attribute("Version") == "1", "Unsupported provider.");
        Require(Leaf(info, "EncryptScope") == "Partial" && Leaf(info, "DecryptSeedLoc") == "/" + SeedPath &&
                Leaf(info, "EntriesMapLoc") == "/" + MapPath, "Unsupported encryption locations or scope.");
        var seed = Parse(Get(entries, SeedPath), settings);
        Require(seed.Name == Ns + "DecyptSeed" && (string?)seed.Attribute("ID") == "1" &&
                (string?)seed.Attribute("EncryptCaseId") == "1.1.1", "Unsupported key descriptor.");
        Shape(seed, ["ID", "EncryptCaseId"], [Ns + "UserInfo", Ns + "ExtendParams"]);
        Shape(Single(seed, Ns + "ExtendParams"), [], []);
        var user = Single(seed, Ns + "UserInfo");
        Shape(user, ["UserName", "UserType"], [Ns + "EncryptedWK", Ns + "IVValue"]);
        Require((string?)user.Attribute("UserType") == "User" && !string.IsNullOrEmpty((string?)user.Attribute("UserName")), "Unsupported user descriptor.");
        var iv = Base64(Leaf(user, "IVValue"), 16);
        var wrapped = Base64(Leaf(user, "EncryptedWK"), 32);
        owned.Add(iv); owned.Add(wrapped);
        var fek = PasswordPrimitives.Cbc(false, wrapped, kek, iv, token); owned.Add(fek);
        Require(fek.Length == 16, "Password or encrypted data is invalid.");
        var mapCipher = Get(entries, MapPath);
        Require(mapCipher.Length <= settings.Metadata + 16, "Encrypted map exceeds metadata budget.");
        var mapBytes = PasswordPrimitives.Cbc(false, mapCipher, fek, iv, token); owned.Add(mapBytes);
        var map = Parse(mapBytes, settings);
        Require(map.Name == Ns + "EncryptEntries" && (string?)map.Attribute("ID") == "1" &&
            (string?)map.Attribute(Profile + "Version") == "1", "Unsupported recovery map.");
        Shape(map, ["ID", Profile + "Version"], [Ns + "EncryptEntry", Profile + "Inventory"], allowRepeated: Ns + "EncryptEntry");
        var mappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var cipherPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in map.Elements(Ns + "EncryptEntry"))
        {
            token.ThrowIfCancellationRequested();
            Shape(item, ["Path", "EPath"], []);
            var path = AbsolutePath(item, "Path"); var cipherPath = AbsolutePath(item, "EPath");
            Require(!Reserved(path) && cipherPath.StartsWith("PasswordCrypto/entry-", StringComparison.Ordinal) &&
                !cipherPath.Substring("PasswordCrypto/".Length).Contains('/') && cipherPath.EndsWith(".dat", StringComparison.Ordinal), "Invalid entry mapping.");
            Require(!entries.ContainsKey(path) && mappings.TryAdd(path, cipherPath) && cipherPaths.Add(cipherPath), "Duplicate or colliding entry mapping.");
        }
        Require(mappings.Count > 0 && mappings.Count <= settings.Load.MaxEntryCount, "Empty or excessive entry mapping.");
        var inventory = Single(map, Profile + "Inventory");
        Shape(inventory, [], [Profile + "Entry"], allowRepeated: Profile + "Entry");
        var original = new Dictionary<string, (long Length, byte[] Hash)>(StringComparer.OrdinalIgnoreCase);
        long total = 0;
        foreach (var item in inventory.Elements())
        {
            token.ThrowIfCancellationRequested();
            Shape(item, ["Path", "Length", "SM3"], []);
            var path = AbsolutePath(item, "Path");
            Require(!Reserved(path) && long.TryParse((string?)item.Attribute("Length"), NumberStyles.None, CultureInfo.InvariantCulture, out var length) &&
                    length <= settings.Load.MaxEntryUncompressedBytes, "Invalid recovery length or path.");
            // Parse only after validation, preserving the compiler's definite assignment guarantees.
            var size = long.Parse(item.Attribute("Length")!.Value, CultureInfo.InvariantCulture);
            total = checked(total + size);
            Require(total <= settings.Load.MaxTotalUncompressedBytes, "Recovery exceeds expanded budget.");
            Require(original.TryAdd(path, (size, Base64(item.Attribute("SM3")!.Value, 32))), "Duplicate recovery path.");
        }
        Require(original.Count <= settings.Load.MaxEntryCount && mappings.Keys.All(original.ContainsKey), "Invalid recovery inventory.");
        var expected = new HashSet<string>(original.Keys.Where(path => !mappings.ContainsKey(path)), StringComparer.OrdinalIgnoreCase);
        expected.UnionWith(cipherPaths); expected.UnionWith([EnvelopePath, SeedPath, MapPath]);
        Require(expected.SetEquals(entries.Keys), "Encrypted package has missing or unexpected entries.");
        foreach (var item in mappings)
        {
            token.ThrowIfCancellationRequested();
            var cipher = Get(entries, item.Value);
            Require(cipher.Length == (original[item.Key].Length / 16 + 1) * 16, "Inconsistent ciphertext length.");
            var plaintext = PasswordPrimitives.Cbc(false, cipher, fek, iv, token); owned.Add(plaintext);
            entries.Remove(item.Value); entries.Add(item.Key, plaintext);
        }
        entries.Remove(EnvelopePath); entries.Remove(SeedPath); entries.Remove(MapPath);
        foreach (var item in original)
        {
            token.ThrowIfCancellationRequested();
            var bytes = Get(entries, item.Key);
            Require(bytes.Length == item.Value.Length && CryptographicOperations.FixedTimeEquals(PasswordPrimitives.Sm3(bytes), item.Value.Hash),
                "Password or damaged recovery data is invalid.");
        }
        ValidateDocument(entries, settings);
        CheckExpanded(entries, settings);
    }

    private static XElement ValidateDocument(Dictionary<string, byte[]> entries, Settings settings)
    {
        var root = Parse(Get(entries, "OFD.xml"), settings);
        Require(root.Name.LocalName == "OFD" && (root.Name.NamespaceName == Ns.NamespaceName ||
            root.Name.NamespaceName == "http://www.ofdspec.org/2016"), "Unrecognized OFD root.");
        var bodies = root.Elements(root.Name.Namespace + "DocBody").ToArray();
        Require(bodies.Length == 1, "This password profile requires exactly one DocBody.");
        Require(!bodies[0].Elements().Any(item => item.Name.LocalName == "Signatures") &&
            !entries.Keys.Any(path => path.Split('/').Any(part => part.Equals("Signs", StringComparison.OrdinalIgnoreCase)) ||
                path.EndsWith("Signatures.xml", StringComparison.OrdinalIgnoreCase)), "Signed packages require explicit signature cleanup before encryption.");
        var docPath = Reference("OFD.xml", Single(bodies[0], root.Name.Namespace + "DocRoot").Value);
        var document = Parse(Get(entries, docPath), settings);
        Require(document.Name == root.Name.Namespace + "Document", "Unrecognized document root.");
        return root;
    }

    private static HashSet<string> SelectEntries(Dictionary<string, byte[]> entries, XElement root, Settings settings)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in settings.Paths)
        {
            Canonical(path); Get(entries, path); result.Add(entries.Keys.Single(key => key.Equals(path, StringComparison.OrdinalIgnoreCase)));
        }
        if (settings.Pages.Length == 0) return result;
        var docPath = Reference("OFD.xml", Single(Single(root, root.Name.Namespace + "DocBody"), root.Name.Namespace + "DocRoot").Value);
        var doc = Parse(Get(entries, docPath), settings);
        var pages = Single(doc, root.Name.Namespace + "Pages").Elements(root.Name.Namespace + "Page").ToArray();
        Require(pages.Length <= settings.Load.MaxPageCount, "Page count exceeds budget.");
        var pagePaths = pages.Select(page => Reference(docPath, (string?)page.Attribute("BaseLoc") ?? throw new InvalidDataException("Page BaseLoc missing."))).ToArray();
        var selectedIndices = new HashSet<int>(settings.Pages);
        foreach (var index in selectedIndices)
        {
            if (index < 0 || index >= pages.Length) throw new ArgumentOutOfRangeException(nameof(OfdPasswordOptions.PageIndices));
            Require(!pagePaths.Where((_, other) => !selectedIndices.Contains(other)).Contains(pagePaths[index], StringComparer.OrdinalIgnoreCase),
                "Selected and unselected pages share a content entry; select explicit entries instead.");
            Get(entries, pagePaths[index]);
            result.Add(entries.Keys.Single(key => key.Equals(pagePaths[index], StringComparison.OrdinalIgnoreCase)));
        }
        return result;
    }

    private static string Reference(string container, string value)
    {
        Require(!string.IsNullOrWhiteSpace(value) && !value.Contains('\\') && !value.Contains(':') && !value.Any(char.IsControl), "Unsafe package reference.");
        return Canonical(OfdPackagePath.Resolve(container, value));
    }

    private static string Canonical(string path)
    {
        Require(!string.IsNullOrWhiteSpace(path) && !path.StartsWith('/') && !path.Contains('\\') && !path.Contains(':') && !path.Any(char.IsControl) &&
            path.Split('/').All(part => part.Length > 0 && part != "." && part != "..") && OfdPackagePath.Resolve("OFD.xml", "/" + path) == path,
            "Unsafe or noncanonical package path.");
        return path;
    }

    private static string AbsolutePath(XElement item, string attribute)
    {
        var value = (string?)item.Attribute(attribute);
        Require(value is not null && value.StartsWith('/'), "Map requires absolute package paths.");
        return Canonical(value!.Substring(1));
    }

    private static bool Reserved(string path) => path.Equals(EnvelopePath, StringComparison.OrdinalIgnoreCase) ||
        path.Equals("PasswordCrypto", StringComparison.OrdinalIgnoreCase) || path.StartsWith("PasswordCrypto/", StringComparison.OrdinalIgnoreCase) ||
        path.Equals("decryptseed.dat", StringComparison.OrdinalIgnoreCase) || path.Equals("entriesmap.dat", StringComparison.OrdinalIgnoreCase);

    private static byte[] Get(Dictionary<string, byte[]> entries, string path) => entries.TryGetValue(path, out var bytes)
        ? bytes : throw new InvalidDataException($"Required package entry is missing: {path}");

    private static byte[] Base64(string value, int length)
    {
        try
        {
            Require(value.Length <= (length + 2) / 3 * 4, "Invalid encoded field length.");
            var bytes = Convert.FromBase64String(value);
            Require(bytes.Length == length, "Invalid decoded field length.");
            return bytes;
        }
        catch (FormatException ex) { throw new InvalidDataException("Invalid base64 field.", ex); }
    }

    private static XElement Parse(byte[] bytes, Settings settings)
    {
        Require(bytes.Length <= settings.Metadata, "XML exceeds metadata budget.");
        try
        {
            using var stream = new MemoryStream(bytes, false);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = settings.Metadata,
                IgnoreComments = true, IgnoreProcessingInstructions = true
            });
            return XDocument.Load(reader).Root ?? throw new InvalidDataException("Missing XML root.");
        }
        catch (XmlException ex) { throw new InvalidDataException("Invalid XML or encrypted data.", ex); }
    }

    internal static byte[] XmlBytes(XElement root, int maximumBytes, OwnedBuffers owned)
    {
        var bytes = new UTF8Encoding(false).GetBytes(root.ToString(SaveOptions.DisableFormatting));
        owned.Add(bytes);
        Require(bytes.Length <= maximumBytes, "Generated metadata exceeds budget.");
        return bytes;
    }

    private static XElement Single(XElement root, XName name)
    {
        var matches = root.Elements(name).ToArray();
        Require(matches.Length == 1, $"Expected one {name.LocalName} node.");
        return matches[0];
    }

    private static string Leaf(XElement root, string name)
    {
        var node = Single(root, Ns + name); Shape(node, [], []);
        return node.Value;
    }

    private static void Shape(XElement node, XName[] attributes, XName[] children, XName? allowRepeated = null)
    {
        Require(node.Attributes().Where(attr => !attr.IsNamespaceDeclaration).All(attr => attributes.Contains(attr.Name)) &&
            attributes.All(name => node.Attribute(name) is not null) && node.Elements().All(child => children.Contains(child.Name)) &&
            children.Where(name => name != allowRepeated).All(name => node.Elements(name).Count() == 1) &&
            (!node.HasElements || !node.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))), "Unsupported or ambiguous XML structure.");
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidDataException(message);
    }

    private static void CheckExpanded(Dictionary<string, byte[]> entries, Settings settings)
    {
        Require(entries.Count <= settings.Load.MaxEntryCount && entries.Values.All(bytes => bytes.LongLength <= settings.Load.MaxEntryUncompressedBytes) &&
            entries.Values.Sum(bytes => bytes.LongLength) <= settings.Load.MaxTotalUncompressedBytes, "Result exceeds package expansion budget.");
    }

    private static async Task<byte[]> Pack(Dictionary<string, byte[]> entries, Settings settings, CancellationToken token)
    {
        using var output = new BoundedMemoryStream(Math.Min(settings.Output, settings.Load.MaxInputBytes));
        try
        {
            using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
                foreach (var entry in entries.OrderBy(pair => pair.Key, StringComparer.Ordinal))
                {
                    token.ThrowIfCancellationRequested();
                    var target = zip.CreateEntry(entry.Key, CompressionLevel.NoCompression);
                    target.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
                    using var stream = target.Open();
                    await stream.WriteAsync(entry.Value, token).ConfigureAwait(false);
                }
            token.ThrowIfCancellationRequested();
            return output.ToArray();
        }
        finally { CryptographicOperations.ZeroMemory(output.GetBuffer()); }
    }

    private static async Task TransformFileAsync(string sourcePath, string destinationPath, string password, OfdPasswordOptions options, bool encrypt, CancellationToken token)
    {
        var sourceFull = Path.GetFullPath(sourcePath); var destinationFull = Path.GetFullPath(destinationPath);
        if (sourceFull.Equals(destinationFull, StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Input and output files must differ.");
        byte[]? result = null; string? temporary = null;
        try
        {
            await using (var input = new FileStream(sourceFull, FileMode.Open, FileAccess.Read, FileShare.Read))
                result = await TransformAsync(input, password, options, encrypt, token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            temporary = Path.Combine(Path.GetDirectoryName(destinationFull)!, ".ofd-password-" + Guid.NewGuid().ToString("N") + ".tmp");
            var fileOptions = new FileStreamOptions { Mode = FileMode.CreateNew, Access = FileAccess.Write, Share = FileShare.None, Options = FileOptions.Asynchronous };
            if (!OperatingSystem.IsWindows()) fileOptions.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            await using (var output = new FileStream(temporary, fileOptions))
            {
                await output.WriteAsync(result, token).ConfigureAwait(false);
                await output.FlushAsync(token).ConfigureAwait(false);
            }
            token.ThrowIfCancellationRequested();
            File.Move(temporary, destinationFull, overwrite: true);
            temporary = null;
        }
        finally
        {
            if (result is not null) CryptographicOperations.ZeroMemory(result);
            if (temporary is not null && File.Exists(temporary)) File.Delete(temporary);
        }
    }

    internal sealed class BoundedMemoryStream(long maximum) : MemoryStream
    {
        public override int Capacity
        {
            get => base.Capacity;
            set
            {
                var previous = GetBuffer();
                base.Capacity = value;
                if (!ReferenceEquals(previous, GetBuffer())) CryptographicOperations.ZeroMemory(previous);
            }
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) CryptographicOperations.ZeroMemory(GetBuffer());
            base.Dispose(disposing);
        }
        private void Check(int count) => Require(Position <= maximum - count, "Result exceeds output ZIP budget.");
        public override void Write(byte[] buffer, int offset, int count) { Check(count); base.Write(buffer, offset, count); }
        public override void Write(ReadOnlySpan<byte> buffer) { Check(buffer.Length); base.Write(buffer); }
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token) { Check(count); return base.WriteAsync(buffer, offset, count, token); }
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default) { Check(buffer.Length); return base.WriteAsync(buffer, token); }
        public override void WriteByte(byte value) { Check(1); base.WriteByte(value); }
        public override void SetLength(long value) { Require(value <= maximum, "Result exceeds output ZIP budget."); base.SetLength(value); }
    }

    private sealed class Settings
    {
        internal OfdPackageLoadOptions Load { get; }
        internal long Output { get; }
        internal int Metadata { get; }
        internal string[] Paths { get; }
        internal int[] Pages { get; }
        internal string UserName { get; }
        internal Settings(OfdPasswordOptions options)
        {
            ArgumentNullException.ThrowIfNull(options); ArgumentNullException.ThrowIfNull(options.LoadOptions);
            if (options.MaxOutputBytes <= 0 || options.MaxOutputBytes > int.MaxValue || options.MaxMetadataBytes <= 0 ||
                options.MaxMetadataBytes > 16 * 1024 * 1024 || string.IsNullOrWhiteSpace(options.UserName) || options.UserName.Length > 256)
                throw new ArgumentOutOfRangeException(nameof(options), "Invalid password profile budgets or user name.");
            var input = options.LoadOptions;
            if (input.MaxInputBytes <= 0 || input.MaxInputBytes > int.MaxValue || input.MaxEntryCount <= 0)
                throw new ArgumentOutOfRangeException(nameof(options), "Snapshot input must fit an in-memory array and entry count must be positive.");
            if (input.MaxCompressionRatio < 1 || double.IsNaN(input.MaxCompressionRatio) || double.IsInfinity(input.MaxCompressionRatio))
                throw new ArgumentOutOfRangeException(nameof(options), "This uncompressed password profile requires a finite compression ratio budget of at least 1.");
            Load = new OfdPackageLoadOptions
            {
                MaxInputBytes = input.MaxInputBytes, MaxEntryCount = input.MaxEntryCount, MaxPageCount = input.MaxPageCount,
                MaxEntryUncompressedBytes = input.MaxEntryUncompressedBytes, MaxTotalUncompressedBytes = input.MaxTotalUncompressedBytes,
                MaxCompressionRatio = input.MaxCompressionRatio, MaxAnnotationObjectCount = input.MaxAnnotationObjectCount,
                MaxAnnotationXmlBytes = input.MaxAnnotationXmlBytes
            };
            Output = options.MaxOutputBytes; Metadata = options.MaxMetadataBytes; UserName = options.UserName;
            Paths = options.EntryNames.ToArray(); Pages = options.PageIndices.ToArray();
            if (Paths.Distinct(StringComparer.OrdinalIgnoreCase).Count() != Paths.Length || Pages.Distinct().Count() != Pages.Length)
                throw new ArgumentException("Duplicate selections are not allowed.", nameof(options));
            if (Paths.Length > Load.MaxEntryCount || Pages.Length > Load.MaxPageCount) throw new ArgumentOutOfRangeException(nameof(options), "Selection exceeds budget.");
        }
    }
}
