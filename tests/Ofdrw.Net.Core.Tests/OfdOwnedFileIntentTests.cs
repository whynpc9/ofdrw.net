using System.Runtime.InteropServices;
using Ofdrw.Net.Core.Models;
using Xunit.Abstractions;

namespace Ofdrw.Net.Core.Tests;

public sealed class OfdOwnedFileIntentTests(ITestOutputHelper output)
{
    [Fact]
    public void ActualFreshTicketCaseVariantsFollowPlatformPolicy()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "ofdrw-core-case-intent-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var ticket = OfdOwnedSaveAuthority.BeginFreshJob(new OfdModelLimits(), default);
        try
        {
            var input = Path.Combine(directory, "Input.ofd"); var target = Path.Combine(directory, "input.ofd");
            File.WriteAllText(input, "original input sentinel");
            var aliasOnVolume = File.Exists(target);
            var insensitive = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX);
            output.WriteLine($"Actual OS: {RuntimeInformation.OSDescription}; case names alias on volume: {aliasOnVolume}; insensitive policy: {insensitive}");
            if (insensitive)
            {
                Assert.Throws<InvalidDataException>(() => OfdOwnedSaveAuthority.BindFileIntent(ticket, input, target, default));
                AssertTerminal(ticket);
                Assert.Throws<InvalidOperationException>(() => OfdOwnedSaveAuthority.BindFileIntent(ticket, input, target, default));
            }
            else
            {
                Assert.False(aliasOnVolume); // prove the independent-file host prerequisite rather than simulate it
                File.WriteAllText(target, "independent target sentinel");
                OfdOwnedSaveAuthority.BindFileIntent(ticket, input, target, default);
                Assert.Equal((int)OfdOwnedSavePhase.Bootstrap, OfdOwnedSaveAuthority.CaptureStatistics(ticket).PhaseCode);
                Assert.Equal("independent target sentinel", File.ReadAllText(target));
                Assert.Throws<InvalidOperationException>(() => OfdOwnedSaveAuthority.BindFileIntent(ticket, input, target, default));
                AssertTerminal(ticket); // successful binding is still once-only
            }
            Assert.Equal("original input sentinel", File.ReadAllText(input));
        }
        finally { OfdOwnedSaveAuthority.Fail(ticket); Directory.Delete(directory, true); }
    }

    [Fact]
    public void ActualFreshTicketSameCanonicalPathRejectsAndRemainsTerminal()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "ofdrw-core-same-intent-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var ticket = OfdOwnedSaveAuthority.BeginFreshJob(new OfdModelLimits(), default);
        try
        {
            var input = Path.Combine(directory, "input.ofd"); File.WriteAllText(input, "original input sentinel");
            Assert.Throws<InvalidDataException>(() => OfdOwnedSaveAuthority.BindFileIntent(ticket, input, input, default));
            AssertTerminal(ticket);
            Assert.Throws<InvalidOperationException>(() => OfdOwnedSaveAuthority.BindFileIntent(ticket, input, Path.Combine(directory, "output.ofd"), default));
            Assert.Equal("original input sentinel", File.ReadAllText(input));
        }
        finally { OfdOwnedSaveAuthority.Fail(ticket); Directory.Delete(directory, true); }
    }
    private static void AssertTerminal(OfdOwnedJobTicket ticket)
    {
        var actual = OfdOwnedSaveAuthority.CaptureStatistics(ticket);
        Assert.True(actual.Available); Assert.Equal((int)OfdOwnedSavePhase.Failed, actual.PhaseCode);
    }
}
