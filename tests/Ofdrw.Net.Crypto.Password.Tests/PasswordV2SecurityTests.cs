using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Crypto.Password.Internal;
using Org.BouncyCastle.Crypto.Generators;

namespace Ofdrw.Net.Crypto.Password.Tests;

public sealed partial class PasswordEnvelopeTests
{
    [Theory]
    [InlineData("12345678", "abdfba6fc9f6641b83bf0c2a9c7cc696cb5b54bb425c01040f75186441fb31159ec392ab1545f70d284552a9993ebabf")]
    [InlineData(Password, "11f8415821a0272b083b8ce7505f6245ab1a4086a3bbf877e898539f3b0fadf0043b2a8c9bdf7ea172b5b4952363aa76")]
    public void V2ScryptMatchesIndependentPythonOpenSsl(string password, string expected)
    {
        var salt = Enumerable.Range(0, 32).Select(n => (byte)n).ToArray();
        Assert.Equal(expected, Hex(PasswordPrimitives.PasswordKey(PasswordPrimitives.PasswordBytes(password), salt, default)));
    }

    [Fact]
    public void PinnedBackendMatchesRfc7914Vector()
    {
        // RFC 7914 section 12: empty P/S, N=16,r=1,p=1,dkLen=64.
        Assert.Equal("77d6576238657b203b19ca42c18a0497f16b4844e3074ae8dfdffa3fede21442fcd0069ded0948f8326a753a0fc81f17e8d3e0fb2e0d3628cf35e20c38d18906",
            Hex(SCrypt.Generate([], [], 16, 1, 1, 64)));
    }

    [Fact]
    public void LogicalFileMacMatchesIndependentPythonFramingVector()
    {
        var entries = new Dictionary<string, byte[]>
        {
            ["Encryptions.xml"] = [0, 1], ["Doc/é.xml"] = [2, 3, 0, 255], ["PasswordCrypto/auth.dat"] = new byte[32]
        };
        var key = Enumerable.Range(0, 32).Select(i => (byte)i).ToArray();
        Assert.Equal("cdb96af8575eed4bfadb7d6dab019271d858f947138ecf079f845777cd09c05d",
            Hex(PasswordPrimitives.Authenticate(entries, "PasswordCrypto/auth.dat", key, default)));
    }

    [Fact]
    public async Task V2SaltAndAllCbcRecordsUseIndependentRandomness()
    {
        var options = Selection(); options.PageIndices.Add(0);
        using var firstSource = new MemoryStream(Zip(Fixture()));
        using var secondSource = new MemoryStream(Zip(Fixture()));
        var a = Unzip(await OfdPasswordEnvelope.EncryptAsync(firstSource, Password, options));
        var b = Unzip(await OfdPasswordEnvelope.EncryptAsync(secondSource, Password, options));
        var salts = new List<byte[]>();
        var ivs = new List<string>();
        foreach (var entries in new[] { a, b })
        {
            var envelope = XElement.Parse(Encoding.UTF8.GetString(entries["Encryptions.xml"]));
            salts.Add(Convert.FromBase64String(envelope.Element(Profile + "Kdf")!.Attribute("Salt")!.Value));
            var seed = XElement.Parse(Encoding.UTF8.GetString(entries["PasswordCrypto/decryptseed.dat"]));
            ivs.Add(seed.Descendants(Ns + "IVValue").Single().Value);
            foreach (var path in new[] { "PasswordCrypto/entry-00001.dat", "PasswordCrypto/entry-00002.dat", "PasswordCrypto/entriesmap.dat" })
                ivs.Add(Convert.ToBase64String(entries[path][..16]));
            AssertEntries(Fixture(), Unzip(await OfdPasswordEnvelope.DecryptAsync(new MemoryStream(Zip(entries)), Password)));
        }
        Assert.Equal(32, salts[0].Length); Assert.NotEqual(salts[0], salts[1]);
        Assert.Equal(ivs.Count, ivs.Distinct().Count());
        Assert.NotEqual(a["PasswordCrypto/auth.dat"], b["PasswordCrypto/auth.dat"]);
    }

    [Theory]
    [InlineData("N", "1")]
    [InlineData("N", "9999999999999999999999999999999999")]
    [InlineData("N", "+131072")]
    [InlineData("N", "0131072")]
    [InlineData("N", " 131072")]
    [InlineData("r", "0")]
    [InlineData("r", "2147483647")]
    [InlineData("p", "2147483647")]
    [InlineData("dkLen", "2147483647")]
    [InlineData("Algorithm", "SM3")]
    [InlineData("Salt", "!")]
    public async Task UntrustedKdfParametersRejectBeforeWaitingForKdf(string field, string value)
    {
        var entries = Unzip(await EncryptFixture());
        MutateXml(entries, "Encryptions.xml", root => root.Element(Profile + "Kdf")!.SetAttributeValue(field, value));
        await PasswordPrimitives.KdfGate.WaitAsync();
        try
        {
            using var cancel = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.DecryptAsync(new MemoryStream(Zip(entries)), Password, cancellationToken: cancel.Token));
        }
        finally { PasswordPrimitives.KdfGate.Release(); }
    }

    [Theory]
    [InlineData("v1")]
    [InlineData("missing-version")]
    [InlineData("foreign-kdf")]
    [InlineData("duplicate-kdf")]
    [InlineData("provider-v1")]
    [InlineData("short-salt")]
    public async Task VersionAndDescriptorConfusionRejectBeforeKdf(string kind)
    {
        var entries = Unzip(await EncryptFixture());
        MutateXml(entries, "Encryptions.xml", root =>
        {
            if (kind == "v1") { root.Attribute(Profile + "Version")!.Remove(); root.SetAttributeValue(XName.Get("Version", "urn:ofdrw-net:password-profile:1"), "1"); }
            if (kind == "missing-version") root.Attribute(Profile + "Version")!.Remove();
            if (kind == "foreign-kdf") root.Element(Profile + "Kdf")!.Name = XName.Get("Kdf", "urn:foreign");
            if (kind == "duplicate-kdf") root.Add(new XElement(root.Element(Profile + "Kdf")!));
            if (kind == "provider-v1") root.Descendants(Ns + "Provider").Single().SetAttributeValue("Version", "1");
            if (kind == "short-salt") root.Element(Profile + "Kdf")!.SetAttributeValue("Salt", Convert.ToBase64String(new byte[31]));
        });
        await PasswordPrimitives.KdfGate.WaitAsync();
        try
        {
            using var cancel = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.DecryptAsync(new MemoryStream(Zip(entries)), Password, cancellationToken: cancel.Token));
        }
        finally { PasswordPrimitives.KdfGate.Release(); }
    }

    [Theory]
    [InlineData("metadata-bytes")]
    [InlineData("username")]
    [InlineData("wrap-iv")]
    [InlineData("map-iv")]
    [InlineData("tag")]
    [InlineData("auth-case")]
    [InlineData("name-case")]
    [InlineData("path")]
    [InlineData("salt")]
    public async Task ExactRawMetadataNamesAndRecordsAreAuthenticated(string kind)
    {
        var entries = Unzip(await EncryptFixture());
        if (kind == "metadata-bytes") entries["Encryptions.xml"] = Encoding.UTF8.GetBytes("<!--synthetic edit-->" + Encoding.UTF8.GetString(entries["Encryptions.xml"]));
        if (kind == "username") MutateXml(entries, "PasswordCrypto/decryptseed.dat", root => root.Descendants(Ns + "UserInfo").Single().SetAttributeValue("UserName", "changed"));
        if (kind == "wrap-iv") MutateXml(entries, "PasswordCrypto/decryptseed.dat", root => root.Descendants(Ns + "IVValue").Single().Value = Convert.ToBase64String(new byte[16]));
        if (kind == "map-iv") entries["PasswordCrypto/entriesmap.dat"][0] ^= 1;
        if (kind == "tag") entries["PasswordCrypto/auth.dat"][0] ^= 1;
        if (kind == "auth-case" || kind == "name-case" || kind == "path")
        {
            var path = kind == "auth-case" ? "PasswordCrypto/auth.dat" : Page0;
            var bytes = entries[path]; entries.Remove(path);
            entries.Add(kind == "path" ? "Doc_0/renamed.xml" : path.ToUpperInvariant(), bytes);
        }
        if (kind == "salt") MutateXml(entries, "Encryptions.xml", root => root.Element(Profile + "Kdf")!.SetAttributeValue("Salt", Convert.ToBase64String(new byte[32])));
        var ex = await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.DecryptAsync(new MemoryStream(Zip(entries)), Password));
        if (kind != "auth-case") Assert.Contains("authenticated", ex.Message);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task KdfBudgetsRefuseRatherThanDowngrade(bool memory)
    {
        var options = Selection();
        if (memory) options.MaxKdfMemoryBytes = PasswordPrimitives.KdfMemoryBudget - 1;
        else options.MaxKdfWorkUnits = PasswordPrimitives.KdfWorkUnits - 1;
        var encrypted = await EncryptFixture();
        await PasswordPrimitives.KdfGate.WaitAsync();
        try
        {
            using var cancel = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.EncryptAsync(new MemoryStream(Zip(Fixture())), Password, options, cancel.Token));
            await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.DecryptAsync(new MemoryStream(encrypted), Password, options, cancel.Token));
        }
        finally { PasswordPrimitives.KdfGate.Release(); }
    }

    [Fact]
    public async Task QueuedKdfCancellationReturnsNoOutputAndGateRecovers()
    {
        var encrypted = await EncryptFixture();
        await PasswordPrimitives.KdfGate.WaitAsync();
        try
        {
            using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(50));
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfdPasswordEnvelope.DecryptAsync(new MemoryStream(encrypted), Password, cancellationToken: cancellation.Token));
        }
        finally { PasswordPrimitives.KdfGate.Release(); }
        AssertEntries(Fixture(), Unzip(await OfdPasswordEnvelope.DecryptAsync(new MemoryStream(encrypted), Password)));
    }

    [Fact]
    public async Task RunningKdfCancellationWaitsForBackendAndPreservesOutput()
    {
        var directory = TempDirectory();
        try
        {
            var input = Path.Combine(directory, "encrypted.ofd"); var output = Path.Combine(directory, "output.ofd");
            await File.WriteAllBytesAsync(input, await EncryptFixture()); var sentinel = new byte[] { 12, 34, 56 };
            await File.WriteAllBytesAsync(output, sentinel);
            using var cancel = new CancellationTokenSource();
            // Cancel on a separate thread only once the production lease is held.
            var cancelTask = Task.Run(async () =>
            {
                while (PasswordPrimitives.KdfGate.CurrentCount != 0) await Task.Delay(1);
                await Task.Delay(30);
                cancel.Cancel();
            });
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfdPasswordEnvelope.DecryptFileAsync(input, output, Password, cancellationToken: cancel.Token));
            await cancelTask;
            Assert.Equal(1, PasswordPrimitives.KdfGate.CurrentCount);
            Assert.Equal(sentinel, await File.ReadAllBytesAsync(output));
            Assert.Empty(Directory.GetFiles(directory, ".ofd-password-*.tmp"));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task PasswordWhitespaceAndUnicodeAreNotReinterpreted()
    {
        var password = " e\u0301 🔒 ";
        var encrypted = await OfdPasswordEnvelope.EncryptAsync(new MemoryStream(Zip(Fixture())), password, Selection());
        AssertEntries(Fixture(), Unzip(await OfdPasswordEnvelope.DecryptAsync(new MemoryStream(encrypted), password)));
        foreach (var other in new[] { password.Trim(), password.Normalize() })
            await Assert.ThrowsAsync<InvalidDataException>(() => OfdPasswordEnvelope.DecryptAsync(new MemoryStream(encrypted), other));
    }
}
