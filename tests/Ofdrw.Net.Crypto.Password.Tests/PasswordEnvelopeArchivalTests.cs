using System.Text;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Crypto.Password.Tests;

public sealed partial class PasswordEnvelopeTests
{
    // An encrypted OFD-A contradicts GB/T 42133 clause 6.16: encryption refuses archival originals,
    // while an explicit ordinary declaration keeps working.
    [Theory]
    [InlineData("OFD-A", true)]
    [InlineData("OFD", false)]
    public async Task Encryption_RejectsArchivalOriginalAndLeavesNoOutput(string docType, bool rejected)
    {
        var original = Fixture();
        original["OFD.xml"] = Encoding.UTF8.GetBytes($"<OFD xmlns='{Ns}' DocType='{docType}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        var directory = TempDirectory();
        try
        {
            var input = Path.Combine(directory, "plain.ofd"); var output = Path.Combine(directory, "enc.ofd");
            await File.WriteAllBytesAsync(input, Zip(original));
            var before = await File.ReadAllBytesAsync(input);
            if (!rejected)
            {
                await OfdPasswordEnvelope.EncryptFileAsync(input, output, Password, Selection());
                Assert.True(File.Exists(output));
                return;
            }
            using (var stream = new MemoryStream(before))
            {
                var error = await Assert.ThrowsAsync<OfdArchivalOriginalException>(() => OfdPasswordEnvelope.EncryptAsync(stream, Password, Selection()));
                Assert.Equal("OfdPasswordEnvelope.EncryptAsync", error.Operation);
                Assert.Contains("OFD-A", error.Message);
            }
            await Assert.ThrowsAsync<OfdArchivalOriginalException>(() => OfdPasswordEnvelope.EncryptFileAsync(input, output, Password, Selection()));
            Assert.False(File.Exists(output));
            Assert.Equal(new[] { input }, Directory.GetFiles(directory));
            Assert.Equal(before, await File.ReadAllBytesAsync(input));
            // Decryption is not a content change and is not guarded.
        }
        finally { Directory.Delete(directory, true); }
    }
}
