if (args.Length < 4 || args[1] != "--")
{
    Console.Error.WriteLine("Usage: Ofdrw.Net.Cli.Tests.dll REPORT.json -- INPUT.ofd OUTPUT.ofd --find TEXT --replace TEXT --trusted-font SHA:BYTES [--pages LIST]");
    return 2;
}
using var shutdown = new CancellationTokenSource();
Console.CancelKeyPress += (_, signal) => { signal.Cancel = true; shutdown.Cancel(); };
var result = await Cli.CaptureOwnedDispatchForDiagnosticsAsync(args[2..], shutdown.Token).ConfigureAwait(false);
await File.WriteAllTextAsync(args[0], result.Json, CancellationToken.None).ConfigureAwait(false);
Console.WriteLine(Path.GetFullPath(args[0]));
return result.ExitCode;
