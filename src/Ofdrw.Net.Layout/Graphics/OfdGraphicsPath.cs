using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Layout.Graphics;

/// <summary>OFD path filling rule.</summary>
public enum OfdFillRule
{
    /// <summary>Nonzero winding number.</summary>
    NonZero,
    /// <summary>Even-odd winding parity.</summary>
    EvenOdd
}

/// <summary>A bounded, mutable path builder in user-space millimeters. Draw/clip operations snapshot it.</summary>
public sealed class OfdGraphicsPath
{
    private readonly List<(string Command, double[] Values)> _commands = new();
    private bool _figure;
    /// <summary>Creates a path. Only absolute M/L/B/Q/C commands are emitted; cubic is B, close is C in OFD.</summary>
    public OfdGraphicsPath(OfdFillRule fillRule = OfdFillRule.NonZero, int maxCommands = 100_000)
    {
        if (fillRule is not (OfdFillRule.NonZero or OfdFillRule.EvenOdd)) throw new ArgumentOutOfRangeException(nameof(fillRule));
        if (maxCommands <= 0) throw new ArgumentOutOfRangeException(nameof(maxCommands));
        FillRule = fillRule; MaxCommands = maxCommands;
    }
    /// <summary>Fill rule used by drawing and clipping.</summary>
    public OfdFillRule FillRule { get; }
    /// <summary>Maximum commands accepted by this builder.</summary>
    public int MaxCommands { get; }
    /// <summary>Number of commands currently stored.</summary>
    public int CommandCount => _commands.Count;
    /// <summary>Starts a new figure.</summary>
    public OfdGraphicsPath MoveTo(double x, double y) { Add("M", x, y); _figure = true; return this; }
    /// <summary>Adds a line from the current point.</summary>
    public OfdGraphicsPath LineTo(double x, double y) { RequireFigure(); Add("L", x, y); return this; }
    /// <summary>Adds a cubic Bezier curve from the current point.</summary>
    public OfdGraphicsPath BezierTo(double x1, double y1, double x2, double y2, double x, double y)
    { RequireFigure(); Add("B", x1, y1, x2, y2, x, y); return this; }
    /// <summary>Adds a quadratic Bezier curve from the current point.</summary>
    public OfdGraphicsPath QuadraticTo(double cx, double cy, double x, double y)
    { RequireFigure(); Add("Q", cx, cy, x, y); return this; }
    /// <summary>Closes a figure. Start another with MoveTo before adding further segments.</summary>
    public OfdGraphicsPath Close() { RequireFigure(); Add("C"); _figure = false; return this; }
    /// <summary>Adds a positive-size closed rectangle atomically.</summary>
    public OfdGraphicsPath AddRectangle(double x, double y, double width, double height)
    {
        GraphicsValidation.Finite(x, y, x + width, y + height);
        GraphicsValidation.Positive(width); GraphicsValidation.Positive(height);
        if (_commands.Count > MaxCommands - 5) throw new InvalidOperationException("Path command budget exceeded.");
        return MoveTo(x, y).LineTo(x + width, y).LineTo(x + width, y + height).LineTo(x, y + height).Close();
    }
    private void Add(string command, params double[] values)
    {
        GraphicsValidation.Finite(values);
        if (_commands.Count >= MaxCommands) throw new InvalidOperationException("Path command budget exceeded.");
        _commands.Add((command, values));
    }
    private void RequireFigure() { if (!_figure) throw new InvalidOperationException("Start a path figure with MoveTo."); }
    internal PathSnapshot Snapshot(int maximum, CancellationToken token)
    {
        if (_commands.Count > maximum) throw new InvalidOperationException("Graphics path command budget exceeded.");
        if (!_commands.Any(command => command.Command is "L" or "B" or "Q")) throw new ArgumentException("A path must contain a drawable segment.");
        var values = new List<(string Command, double[] Values)>(_commands.Count);
        foreach (var command in _commands) { token.ThrowIfCancellationRequested(); values.Add((command.Command, (double[])command.Values.Clone())); }
        return new PathSnapshot(values, FillRule);
    }
}

internal sealed class PathSnapshot
{
    internal PathSnapshot(List<(string Command, double[] Values)> commands, OfdFillRule rule) { Commands = commands; Rule = rule; }
    internal List<(string Command, double[] Values)> Commands { get; }
    internal OfdFillRule Rule { get; }
    internal string Data => string.Join(" ", Commands.Select(command => command.Command + (command.Values.Length == 0 ? "" : " " + string.Join(" ", command.Values.Select(GraphicsValidation.Number)))));
    internal IEnumerable<(double X, double Y)> Points => Commands.SelectMany(command => Enumerable.Range(0, command.Values.Length / 2).Select(i => (command.Values[i * 2], command.Values[i * 2 + 1])));
}
