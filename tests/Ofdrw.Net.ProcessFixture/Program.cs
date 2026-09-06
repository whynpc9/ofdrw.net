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
    case "wait":
        await File.WriteAllTextAsync(args[1], Environment.ProcessId.ToString());
        await Task.Delay(TimeSpan.FromSeconds(30));
        return 0;
    case "tree":
        var start = new ProcessStartInfo(Environment.ProcessPath!) { UseShellExecute = false };
        start.ArgumentList.Add("exec");
        start.ArgumentList.Add("--runtimeconfig");
        start.ArgumentList.Add(args[2]);
        start.ArgumentList.Add(typeof(ProcessFixtureMarker).Assembly.Location);
        start.ArgumentList.Add("wait");
        start.ArgumentList.Add(args[1] + ".child");
        using (var child = Process.Start(start)!)
        {
            await File.WriteAllTextAsync(args[1], Environment.ProcessId.ToString());
            await child.WaitForExitAsync();
        }
        return 0;
    default:
        return 2;
}

public sealed class ProcessFixtureMarker;
