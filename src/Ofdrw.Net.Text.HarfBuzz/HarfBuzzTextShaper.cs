using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using HarfBuzzSharp;
using Ofdrw.Net.Layout.Text;

namespace Ofdrw.Net.Text.HarfBuzz;

/// <summary>The pinned HarfBuzzSharp implementation of the single-run shaping contract.</summary>
public sealed class HarfBuzzTextShaper : ITextShaper
{
    /// <summary>Managed and native assets are pinned together in the adapter package.</summary>
    public const string PackageVersion = "14.2.1.301";
    private const string ExpectedNativeVersion = "14.2.1";

    [DllImport("libHarfBuzzSharp", CallingConvention = CallingConvention.Cdecl, EntryPoint = "hb_version_string")]
    private static extern IntPtr NativeVersionString();

    /// <summary>Shapes one run with the pinned native engine and returns GIDs, positions, and source clusters.</summary>
    public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            var nativeVersion = Marshal.PtrToStringAnsi(NativeVersionString());
            if (nativeVersion != ExpectedNativeVersion)
                throw new TextLayoutException(TextLayoutError.NativeUnavailable,
                    "HarfBuzz native engine version differs from the pinned adapter: " + nativeVersion + ".");
            var engineIdentity = "HarfBuzzSharp/" + PackageVersion + ";libHarfBuzzSharp/" + nativeVersion;
            if (request.Text.Length == 0)
                return new TextShapedRun(request, Array.Empty<TextShapedGlyph>(), engineIdentity);
            using var blob = CreateOwnedBlob(request.Face.GetFaceBytes());
            using var face = new Face(blob, 0);
            if (face.UnitsPerEm != request.Face.UnitsPerEm)
                throw new TextLayoutException(TextLayoutError.InvalidFont, "HarfBuzz and the managed reader disagree about font units.");
            using var font = new Font(face);
            font.SetFunctionsOpenType();
            font.SetScale(request.Face.UnitsPerEm, request.Face.UnitsPerEm);
            using var buffer = new HarfBuzzSharp.Buffer
            {
                Direction = request.Direction == TextDirection.RightToLeft ? Direction.RightToLeft : Direction.LeftToRight,
                Script = Script.Parse(request.Script),
                Language = new Language(request.Language),
                ClusterLevel = ClusterLevel.MonotoneGraphemes,
            };
            buffer.AddUtf16(request.Text);
            var features = request.Features.Select(value => new Feature(
                new Tag(value.Tag[0], value.Tag[1], value.Tag[2], value.Tag[3]), value.Value,
                (uint)value.Start, (uint)(value.Start + value.Length))).ToArray();
            cancellationToken.ThrowIfCancellationRequested();
            font.Shape(buffer, features);
            cancellationToken.ThrowIfCancellationRequested();
            var infos = buffer.GlyphInfos;
            var positions = buffer.GlyphPositions;
            if (infos.Length != positions.Length || infos.Length > request.Text.Length * 4L + 16)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Native shaping exceeded the glyph expansion budget.");
            var boundaries = new SortedSet<int> { request.Text.Length };
            foreach (var info in infos)
            {
                if (info.Cluster >= (uint)request.Text.Length ||
                    info.Cluster > 0 && char.IsLowSurrogate(request.Text[(int)info.Cluster]))
                    throw new TextLayoutException(TextLayoutError.InvalidText, "Native shaping returned an invalid UTF-16 cluster.");
                boundaries.Add((int)info.Cluster);
                if (info.Codepoint == 0)
                    throw new TextLayoutException(TextLayoutError.MissingGlyph,
                        "Selected face has an unmapped glyph at source UTF-16 offset " + (request.SourceStart + info.Cluster) + ".",
                        sourceStart: request.SourceStart + (int)info.Cluster,
                        sourceLength: char.IsHighSurrogate(request.Text[(int)info.Cluster]) ? 2 : 1);
            }
            var next = new Dictionary<int, int>();
            var previous = -1;
            foreach (var boundary in boundaries)
            {
                if (previous >= 0) next[previous] = boundary;
                previous = boundary;
            }
            var glyphs = new List<TextShapedGlyph>(infos.Length);
            for (var i = 0; i < infos.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var info = infos[i]; var position = positions[i];
                var cluster = (int)info.Cluster;
                TextInkBounds? ink = null;
                if (font.TryGetGlyphExtents(info.Codepoint, out var extents) &&
                    extents.Width > 0 && extents.Height < 0)
                    ink = new TextInkBounds(extents.XBearing, extents.YBearing + extents.Height,
                        extents.Width, -extents.Height);
                glyphs.Add(new TextShapedGlyph(info.Codepoint, request.SourceStart + cluster,
                    next[cluster] - cluster, position.XAdvance, position.YAdvance,
                    position.XOffset, position.YOffset, ink,
                    (info.GlyphFlags & GlyphFlags.UnsafeToBreak) != 0));
            }
            return new TextShapedRun(request, glyphs, engineIdentity);
        }
        catch (TextLayoutException) { throw; }
        catch (Exception exception) when (exception is DllNotFoundException || exception is EntryPointNotFoundException ||
            exception is BadImageFormatException || exception is TypeInitializationException)
        {
            throw new TextLayoutException(TextLayoutError.NativeUnavailable,
                "The pinned HarfBuzzSharp native asset is unavailable or incompatible.", exception);
        }
        catch (Exception exception) when (exception is InvalidOperationException || exception is ArgumentException || exception is IOException)
        {
            throw new TextLayoutException(TextLayoutError.InvalidFont, "HarfBuzz rejected the selected face or run.", exception);
        }
    }

    private static Blob CreateOwnedBlob(byte[] faceBytes)
    {
        var pinnedBytes = GCHandle.Alloc(faceBytes, GCHandleType.Pinned);
        try
        {
            // Duplicate copies the bytes into native-owned storage before unpinning the array.
            return new Blob(pinnedBytes.AddrOfPinnedObject(), faceBytes.Length, MemoryMode.Duplicate);
        }
        finally
        {
            pinnedBytes.Free();
        }
    }
}

/// <summary>Default entry point for the new query API. Existing Flow rendering remains in its compatibility mode.</summary>
public static class HarfBuzzTextPrelayout
{
    /// <summary>Creates the new prelayout query with HarfBuzz as its default run shaper.</summary>
    public static TextPrelayoutEngine Create() => new TextPrelayoutEngine(new HarfBuzzTextShaper());

    /// <summary>Creates the bounded mixed Latin/Hebrew paragraph query with the pinned run shaper.</summary>
    public static TextMixedParagraphEngine CreateMixedParagraph() => new TextMixedParagraphEngine(new HarfBuzzTextShaper());

    /// <summary>Creates a whole-run query retaining source-span ownership with the pinned shaper.</summary>
    public static TextSpanShapingEngine CreateSpanShaping() => new TextSpanShapingEngine(new HarfBuzzTextShaper());
}
