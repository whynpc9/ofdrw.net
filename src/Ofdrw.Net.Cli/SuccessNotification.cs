internal static partial class Cli
{
    private static void WriteSuccessNotification(string message)
    {
        // Call only after saving has completed. Notification failure cannot
        // reverse publication or invite a retry of an already committed edit.
        try { Console.WriteLine(message); }
        catch { }
    }
}
