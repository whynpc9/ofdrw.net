using System;
using System.IO;

namespace Ofdrw.Net.Core.IO;

/// <summary>
/// Marks <see cref="InvalidDataException"/> instances raised because a configured resource limit was reached,
/// without changing the exception type callers already observe.
/// </summary>
internal static class OfdLimitExceeded
{
    private const string Key = "Ofdrw.Net.LimitExceeded";

    internal static InvalidDataException Create(string message)
    {
        var exception = new InvalidDataException(message);
        exception.Data[Key] = true;
        return exception;
    }

    internal static bool IsLimit(Exception exception) =>
        exception.Data.Contains(Key) && exception.Data[Key] is true;
}
