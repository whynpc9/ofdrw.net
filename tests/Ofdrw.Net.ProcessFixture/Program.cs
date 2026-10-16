using System.Diagnostics;

switch (args[0])
{
    case "flood":
        var line = new string('x', 1024);
        for (var i = 0; i < 1024; i++)
        {
            Console.Out.WriteLine(line);
            Console.Error.WriteLine(line);
        }
        return 7;
    case "arguments":
        foreach (var argument in args.Skip(1)) Console.WriteLine(Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(argument)));
        return 0;
    case "context":
        // The context probe has a UTF-8 protocol independent of the Windows
        // console code page; its caller selects the same decoding explicitly.
        Console.OutputEncoding = new System.Text.UTF8Encoding(false);
        Console.WriteLine(Directory.GetCurrentDirectory());
        Console.WriteLine(Environment.GetEnvironmentVariable("OFDRW_AUDIT024_VALUE"));
        return 0;
    case "echo":
        Console.Out.WriteLine("normal output");
        Console.Error.WriteLine("normal error");
        return 7;
    case "wait":
        await PublishProcessIdAsync(args[1]);
        await Task.Delay(TimeSpan.FromSeconds(30));
        return 0;
    case "stable-output":
    case "parent-exits-first":
    case "tree":
        var start = new ProcessStartInfo(Environment.ProcessPath!) { UseShellExecute = false };
        start.ArgumentList.Add("exec");
        start.ArgumentList.Add("--runtimeconfig");
        start.ArgumentList.Add(args[2]);
        start.ArgumentList.Add(typeof(ProcessFixtureMarker).Assembly.Location);
        start.ArgumentList.Add("wait");
        start.ArgumentList.Add(args[1] + ".child");
        await PublishProcessIdAsync(args[1]);
        using (var child = Process.Start(start)!)
        {
            if (args[0] == "parent-exits-first")
            {
                while (!File.Exists(args[1] + ".child")) await Task.Delay(10);
                return 0;
            }
            if (args[0] == "stable-output") await File.WriteAllTextAsync(args[3], "completed output");
            await child.WaitForExitAsync();
        }
        return 0;
    default:
        return 2;
}

static async Task PublishProcessIdAsync(string path)
{
    // File existence is the fixture readiness signal. Publish only after the
    // writer has closed, so Windows readers cannot race its exclusive handle.
    var pending = path + ".pending";
    await File.WriteAllTextAsync(pending, Environment.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture));
    File.Move(pending, path);
}

public sealed class ProcessFixtureMarker;
