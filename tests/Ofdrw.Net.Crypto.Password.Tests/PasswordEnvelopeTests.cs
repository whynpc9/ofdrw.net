using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Crypto.Password.Internal;
using Ofdrw.Net.Signatures.Verification;
using Org.BouncyCastle.Crypto.Engines;
using Org.BouncyCastle.Crypto.Parameters;

namespace Ofdrw.Net.Crypto.Password.Tests;

public sealed class PasswordEnvelopeTests
{
    private const string Password = "口令🔒";
    private const string Page0 = "Doc_0/Pages/P0/Content.xml";
    private const string Page1 = "Doc_0/Pages/P1/Content.xml";
    private static readonly XNamespace Ns = "http://www.ofdspec.org";
    private static readonly XNamespace Profile = "urn:ofdrw-net:password-profile:1";

    [Theory]
    [InlineData("abc", "66c7f0f462eeedd9d1f2d46bdc10e4e24167c4875cf2f7a2297da02b8f4ba8e0")]
    [InlineData("abcdabcdabcdabcdabcdabcdabcdabcdabcdabcdabcdabcdabcdabcdabcdabcd", "debe9ff92275b8a138604889c18e5a4d6fdb70e5387e5765293dcba39c0c5732")]
    public void Sm3PublishedVectors(string input, string expected) => Assert.Equal(expected, Hex(PasswordPrimitives.Sm3(Encoding.ASCII.GetBytes(input))));

    [Theory]
    [InlineData("12345678", "34d6effbd5bdc7c8020f619bfd8303b5")]
    [InlineData(Password, "786851358d2b990766550167a3f8981c")]
    public void KdfMatchesIndependentOpenSsl(string password, string expected) => Assert.Equal(expected, Hex(PasswordPrimitives.PasswordKey(password)));

    [Fact]
    public void KdfCounterIsBigEndianAndContinuesBeyondFirstBlock()
    {
        var bytes = Encoding.ASCII.GetBytes("12345678");
        var result = PasswordPrimitives.Kdf(bytes, 48);
        Assert.Equal("34d6effbd5bdc7c8020f619bfd8303b5", Hex(result[..16]));
        Assert.Equal("0e4073094460ecd0cbb91abaadc0886c", Hex(result[32..])); // independently computed OpenSSL counter-2 vector
    }

    [Fact]
    public void Sm4PublishedBlockVector()
    {
        var input = Convert.FromHexString("0123456789abcdeffedcba9876543210");
        var cipher = new SM4Engine(); cipher.Init(true, new KeyParameter(input));
        var output = new byte[16]; cipher.ProcessBlock(input, 0, output, 0);
        Assert.Equal("681edf34d206965e86b3e94f536e4246", Hex(output));
    }

    [Theory]
    [InlineData("", "4b910651754b5553f10cfa0c8a09e9e5")]
    [InlineData("abc", "4301693c448c7da7cff13f84690f7dea")]
    [InlineData("中文 ABC\n", "f823ee782b0cb36e4c143e83ac7f64a6")]
    public void CbcPkcs7MatchesIndependentOpenSsl(string input, string expected)
    {
        var key = Convert.FromHexString("0123456789abcdeffedcba9876543210");
        var iv = Convert.FromHexString("000102030405060708090a0b0c0d0e0f");
        var plain = Encoding.UTF8.GetBytes(input);
        var ciphertext = PasswordPrimitives.Cbc(true, plain, key, iv, default);
        Assert.Equal(expected, Hex(ciphertext));
        Assert.Equal(plain, PasswordPrimitives.Cbc(false, ciphertext, key, iv, default));
    }

    [Fact]
    public void WrappedFileKeyMatchesIndependentOpenSsl()
    {
        var fek = Convert.FromHexString("0123456789abcdeffedcba9876543210");
        var iv = Convert.FromHexString("000102030405060708090a0b0c0d0e0f");
        var kek = PasswordPrimitives.PasswordKey("12345678");
        Assert.Equal("033ff79878445981b81938ec7d4f901c58885f048714e56831e5e1c1349f128c", Hex(PasswordPrimitives.Cbc(true, fek, kek, iv, default)));
        var invalidPadding = PasswordPrimitives.Cbc(true, fek, kek, iv, default); invalidPadding[15] ^= 0x10;
        Assert.Throws<InvalidDataException>(() => PasswordPrimitives.Cbc(false, invalidPadding, kek, iv, default));
        Assert.Throws<InvalidDataException>(() => PasswordPrimitives.Cbc(false, invalidPadding[..31], kek, iv, default));
    }

    [Fact]
    public void InvalidPasswordsFail()
    {
        Assert.ThrowsAny<ArgumentException>(() => PasswordPrimitives.PasswordKey(""));
        Assert.ThrowsAny<ArgumentException>(() => PasswordPrimitives.PasswordKey(new string((char)0xd800, 1)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PartialPageAndAllEntriesRestoreExactEntryBytes(bool all)
    {
        var original = Fixture(); var options = new OfdPasswordOptions();
        if (all) foreach (var path in original.Keys) options.EntryNames.Add(path);
        else options.PageIndices.Add(1);
        using var source = new MemoryStream(Zip(original));
        var encrypted = await OfdPasswordEnvelope.EncryptAsync(source, Password, options);
        Assert.True(source.CanRead);
        var stored = Unzip(encrypted);
        Assert.True(stored.ContainsKey("Encryptions.xml"));
        Assert.False(stored.ContainsKey(Page1));
        if (!all) Assert.Equal(original[Page0], stored[Page0]);
        var envelope = XElement.Parse(Encoding.UTF8.GetString(stored["Encryptions.xml"]));
        Assert.Equal("Partial", envelope.Descendants(Ns + "EncryptScope").Single().Value);
        Assert.Equal("1", (string?)envelope.Attribute(Profile + "Version"));
        var seed = XElement.Parse(Encoding.UTF8.GetString(stored["PasswordCrypto/decryptseed.dat"]));
        Assert.Equal(Ns + "DecyptSeed", seed.Name);
        Assert.Equal("1.1.1", (string?)seed.Attribute("EncryptCaseId"));
        Assert.Equal(32, Convert.FromBase64String(seed.Descendants(Ns + "EncryptedWK").Single().Value).Length);
        using var ciphertext = new MemoryStream(encrypted);
        var restored = Unzip(await OfdPasswordEnvelope.DecryptAsync(ciphertext, Password));
        AssertEntries(original, restored);
        Assert.False(OfdCryptographicCapabilities.SupportsGmT0099EncryptionEnvelope);
        Assert.True(OfdPasswordCapabilities.SupportsSelfRoundTripPasswordEnvelope);
        Assert.False(OfdPasswordCapabilities.VendorInteroperabilityVerified);
    }

    [Fact]
    public async Task EachOperationUsesFreshRandomFileKeyAndIv()
    {
        var first = Unzip(await EncryptFixture()); var second = Unzip(await EncryptFixture());
        Assert.NotEqual(first["PasswordCrypto/decryptseed.dat"], second["PasswordCrypto/decryptseed.dat"]);
        Assert.NotEqual(first["PasswordCrypto/entry-00001.dat"], second["PasswordCrypto/entry-00001.dat"]);
    }

    [Theory]
    [InlineData("wrong-password")]
    [InlineData("middle-block")]
    [InlineData("truncated")]
    [InlineData("map-corrupt")]
    [InlineData("missing-cipher")]
    [InlineData("extra-entry")]
    [InlineData("unselected-corrupt")]
    [InlineData("unknown-profile")]
    [InlineData("duplicate-info")]
    [InlineData("malformed-base64")]
    [InlineData("short-iv")]
    [InlineData("certificate")]
    [InlineData("relative")]
    [InlineData("dtd")]
    [InlineData("missing-inventory")]
    [InlineData("map-traversal")]
    [InlineData("collision")]
    [InlineData("duplicate-inventory")]
    [InlineData("unexpected-map-element")]
    public async Task DamagedOrUnsupportedEnvelopesFailBeforePublishing(string kind)
    {
        var encrypted = Unzip(await EncryptFixture());
        string password = Password;
        switch (kind)
        {
            case "wrong-password": password = "wrong"; break;
            case "middle-block": encrypted["PasswordCrypto/entry-00001.dat"][0] ^= 1; break;
            case "truncated": encrypted["PasswordCrypto/entry-00001.dat"] = encrypted["PasswordCrypto/entry-00001.dat"][..^1]; break;
            case "map-corrupt": encrypted["PasswordCrypto/entriesmap.dat"][0] ^= 1; break;
            case "missing-cipher": encrypted.Remove("PasswordCrypto/entry-00001.dat"); break;
            case "extra-entry": encrypted.Add("extra.xml", Encoding.UTF8.GetBytes("extra")); break;
            case "unselected-corrupt": encrypted[Page0][0] ^= 1; break;
            case "unknown-profile": MutateXml(encrypted, "Encryptions.xml", root => root.SetAttributeValue(Profile + "Version", "2")); break;
            case "duplicate-info": MutateXml(encrypted, "Encryptions.xml", root => root.Add(new XElement(root.Elements().Single()))); break;
            case "malformed-base64": MutateXml(encrypted, "PasswordCrypto/decryptseed.dat", root => root.Descendants(Ns + "EncryptedWK").Single().Value = "!"); break;
            case "short-iv": MutateXml(encrypted, "PasswordCrypto/decryptseed.dat", root => root.Descendants(Ns + "IVValue").Single().Value = "AA=="); break;
            case "certificate": MutateXml(encrypted, "PasswordCrypto/decryptseed.dat", root => root.SetAttributeValue("EncryptCaseId", "1.1.2")); break;
            case "relative": MutateXml(encrypted, "Encryptions.xml", root => root.Elements().Single().SetAttributeValue("Relative", "0")); break;
            case "dtd": encrypted["Encryptions.xml"] = Encoding.UTF8.GetBytes("<!DOCTYPE Encryptions [<!ENTITY x 'bad'>]><Encryptions>&x;</Encryptions>"); break;
            default:
                MutateMap(encrypted, root =>
                {
                    var mapping = root.Elements(Ns + "EncryptEntry").Single();
                    if (kind == "missing-inventory") root.Element(Profile + "Inventory")!.Remove();
                    if (kind == "map-traversal") mapping.SetAttributeValue("Path", "/../outside");
                    if (kind == "collision") mapping.SetAttributeValue("Path", "/" + Page0);
                    if (kind == "duplicate-inventory") root.Element(Profile + "Inventory")!.Add(new XElement(root.Element(Profile + "Inventory")!.Elements().First()));
                    if (kind == "unexpected-map-element") root.Add(new XElement(Ns + "Unknown"));
                });
                break;
        }
        using var input = new MemoryStream(Zip(encrypted));
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.DecryptAsync(input, password));
        var directory = TempDirectory();
        try
        {
            var inputPath = Path.Combine(directory, "input.ofd"); var outputPath = Path.Combine(directory, "output.ofd");
            await File.WriteAllBytesAsync(inputPath, Zip(encrypted)); var sentinel = new byte[] { 1, 3, 5, 7 };
            await File.WriteAllBytesAsync(outputPath, sentinel);
            await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.DecryptFileAsync(inputPath, outputPath, password));
            Assert.Equal(sentinel, await File.ReadAllBytesAsync(outputPath));
            File.Delete(outputPath);
            await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.DecryptFileAsync(inputPath, outputPath, password));
            Assert.False(File.Exists(outputPath));
            Assert.Empty(Directory.GetFiles(directory, ".ofd-password-*.tmp"));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task NoEnvelopeAndAlreadyEncryptedAreRejected()
    {
        using var plain = new MemoryStream(Zip(Fixture()));
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.DecryptAsync(plain, Password));
        using var encrypted = new MemoryStream(await EncryptFixture());
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.EncryptAsync(encrypted, Password, Selection()));
    }

    [Theory]
    [InlineData("signature-declaration")]
    [InlineData("signs-payload")]
    [InlineData("multiple-docs")]
    [InlineData("shared-page")]
    [InlineData("unsafe-name")]
    public async Task AmbiguousOrSignedSourceIsRejected(string kind)
    {
        var original = Fixture();
        if (kind == "signature-declaration") MutateXml(original, "OFD.xml", root => root.Element(Ns + "DocBody")!.Add(new XElement(Ns + "Signatures", "/missing")));
        if (kind == "signs-payload") original.Add("Doc_0/Signs/orphan.dat", [1]);
        if (kind == "multiple-docs") MutateXml(original, "OFD.xml", root => root.Add(new XElement(root.Elements().Single())));
        if (kind == "shared-page") MutateXml(original, "Doc_0/Document.xml", root => root.Descendants(Ns + "Page").Last().SetAttributeValue("BaseLoc", "Pages/P0/Content.xml"));
        if (kind == "unsafe-name") original.Add("Doc_0/./bad.xml", [1]);
        using var input = new MemoryStream(Zip(original));
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.EncryptAsync(input, Password, Selection()));
    }

    [Fact]
    public async Task SelectionFailuresAreObservable()
    {
        foreach (var options in new[] { new OfdPasswordOptions(), new OfdPasswordOptions { PageIndices = { -1 } },
                     new OfdPasswordOptions { PageIndices = { 100 } }, new OfdPasswordOptions { PageIndices = { 0, 0 } } })
        {
            using var source = new MemoryStream(Zip(Fixture()));
            await Assert.ThrowsAnyAsync<ArgumentException>(() => OfdPasswordEnvelope.EncryptAsync(source, Password, options));
        }
    }

    [Theory]
    [InlineData("input")]
    [InlineData("entry")]
    [InlineData("total")]
    [InlineData("metadata")]
    [InlineData("output")]
    [InlineData("count")]
    public async Task ProcessingBudgetsFailAtomically(string budget)
    {
        var options = Selection();
        if (budget == "input") options.LoadOptions.MaxInputBytes = 100;
        if (budget == "entry") options.LoadOptions.MaxEntryUncompressedBytes = 100;
        if (budget == "total") options.LoadOptions.MaxTotalUncompressedBytes = 100;
        if (budget == "metadata") options.MaxMetadataBytes = 100;
        if (budget == "output") options.MaxOutputBytes = 100;
        if (budget == "count") options.LoadOptions.MaxEntryCount = Fixture().Count;
        using var source = new MemoryStream(Zip(Fixture()));
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.EncryptAsync(source, Password, options));
    }

    [Fact]
    public async Task CancellationDuringStagingReturnsNoPlaintext()
    {
        var encrypted = await EncryptFixture(); using var cancellation = new CancellationTokenSource();
        using var input = new CancellingStream(encrypted, cancellation);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfdPasswordEnvelope.DecryptAsync(input, Password, cancellationToken: cancellation.Token));
        Assert.True(input.ReadOccurred);
    }

    [Fact]
    public async Task FileCommitAndFailuresCleanTemporaryPlaintext()
    {
        var directory = TempDirectory();
        try
        {
            var input = Path.Combine(directory, "plain.ofd"); var encrypted = Path.Combine(directory, "enc.ofd"); var restored = Path.Combine(directory, "restored.ofd");
            await File.WriteAllBytesAsync(input, Zip(Fixture()));
            await OfdPasswordEnvelope.EncryptFileAsync(input, encrypted, Password, Selection());
            await File.WriteAllBytesAsync(restored, [9, 8]);
            await OfdPasswordEnvelope.DecryptFileAsync(encrypted, restored, Password);
            AssertEntries(Fixture(), Unzip(await File.ReadAllBytesAsync(restored)));
            if (!OperatingSystem.IsWindows()) Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(restored));
            var old = await File.ReadAllBytesAsync(restored);
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfdPasswordEnvelope.DecryptFileAsync(encrypted, restored, Password, cancellationToken: cancelled.Token));
            Assert.Equal(old, await File.ReadAllBytesAsync(restored));
            var blockedDestination = Path.Combine(directory, "destination-directory"); Directory.CreateDirectory(blockedDestination);
            await Assert.ThrowsAnyAsync<IOException>(() => OfdPasswordEnvelope.DecryptFileAsync(encrypted, blockedDestination, Password));
            Assert.Empty(Directory.GetFiles(directory, ".ofd-password-*.tmp"));
            await Assert.ThrowsAsync<ArgumentException>(() => OfdPasswordEnvelope.EncryptFileAsync(input, input, Password, Selection()));
        }
        finally { Directory.Delete(directory, true); }
    }

    private static OfdPasswordOptions Selection() => new() { PageIndices = { 1 } };
    private static async Task<byte[]> EncryptFixture()
    {
        using var input = new MemoryStream(Zip(Fixture()));
        return await OfdPasswordEnvelope.EncryptAsync(input, Password, Selection());
    }
    private static string Hex(byte[] value) => Convert.ToHexString(value).ToLowerInvariant();
    private static string TempDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), "ofd-password-test-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(path); return path;
    }
    private static Dictionary<string, byte[]> Fixture() => new(StringComparer.OrdinalIgnoreCase)
    {
        ["OFD.xml"] = Encoding.UTF8.GetBytes($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
        ["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes($"<Document xmlns='{Ns}'><Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/><Page ID='2' BaseLoc='Pages/P1/Content.xml'/></Pages></Document>"),
        [Page0] = Encoding.UTF8.GetBytes("中文第一页 ABC"),
        [Page1] = Encoding.UTF8.GetBytes(string.Concat(Enumerable.Repeat("第二页 proportional text ABC 0123456789; ", 10))),
        ["custom/unknown.bin"] = [0, 1, 2, 255], ["empty.bin"] = []
    };
    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true)) foreach (var entry in entries)
        {
            using var stream = zip.CreateEntry(entry.Key, CompressionLevel.NoCompression).Open(); stream.Write(entry.Value);
        }
        return output.ToArray();
    }
    private static Dictionary<string, byte[]> Unzip(byte[] bytes)
    {
        using var input = new MemoryStream(bytes); using var zip = new ZipArchive(input);
        return zip.Entries.ToDictionary(entry => entry.FullName, entry => { using var stream = entry.Open(); using var output = new MemoryStream(); stream.CopyTo(output); return output.ToArray(); }, StringComparer.OrdinalIgnoreCase);
    }
    private static void AssertEntries(Dictionary<string, byte[]> expected, Dictionary<string, byte[]> actual)
    {
        Assert.Equal(expected.Keys.Order(), actual.Keys.Order());
        foreach (var entry in expected) Assert.Equal(entry.Value, actual[entry.Key]);
    }
    private static void MutateXml(Dictionary<string, byte[]> entries, string path, Action<XElement> action)
    {
        var root = XElement.Parse(Encoding.UTF8.GetString(entries[path])); action(root); entries[path] = Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting));
    }
    private static void MutateMap(Dictionary<string, byte[]> entries, Action<XElement> action)
    {
        var seed = XElement.Parse(Encoding.UTF8.GetString(entries["PasswordCrypto/decryptseed.dat"]));
        var iv = Convert.FromBase64String(seed.Descendants(Ns + "IVValue").Single().Value);
        var key = PasswordPrimitives.Cbc(false, Convert.FromBase64String(seed.Descendants(Ns + "EncryptedWK").Single().Value), PasswordPrimitives.PasswordKey(Password), iv, default);
        var map = XElement.Parse(Encoding.UTF8.GetString(PasswordPrimitives.Cbc(false, entries["PasswordCrypto/entriesmap.dat"], key, iv, default)));
        action(map);
        entries["PasswordCrypto/entriesmap.dat"] = PasswordPrimitives.Cbc(true, Encoding.UTF8.GetBytes(map.ToString(SaveOptions.DisableFormatting)), key, iv, default);
    }
    private sealed class CancellingStream(byte[] input, CancellationTokenSource cancellation) : MemoryStream(input)
    {
        public bool ReadOccurred { get; private set; }
        public override bool CanSeek => false;
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
        {
            var read = await base.ReadAsync(buffer, offset, count, token); ReadOccurred = true; cancellation.Cancel(); return read;
        }
    }
}
