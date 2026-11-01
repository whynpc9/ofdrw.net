using System;
using System.Collections.Generic;
using System.Globalization;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Threading;
using SkiaSharp;

namespace Ofdrw.Net.Graphics.SkiaSharp;

public sealed partial class SkiaDrawEvent
{
    internal const string FontReadFailureKey = "Ofdrw.Skia.FontReadFailure.v1";

    // Per-factory test seam. Production supplies null: no global read override,
    // retry, alternate face, or native ownership is exposed publicly.
    internal sealed class NativeFontTableAccess
    {
        internal readonly Func<SKTypeface, uint, string, int>? Size;
        internal readonly Func<SKTypeface, uint, int, IntPtr, string, bool>? Byte;
        internal readonly Action? BeforeCreate;
        internal readonly bool? RecordHistory;
        internal readonly Action? BeforeDiagnosticRecord, BeforeDiagnosticSnapshot;
        internal NativeFontTableAccess(Func<SKTypeface, uint, string, int>? size = null,
            Func<SKTypeface, uint, int, IntPtr, string, bool>? readByte = null, Action? beforeCreate = null, bool? recordHistory = null,
            Action? beforeDiagnosticRecord = null, Action? beforeDiagnosticSnapshot = null)
        {
            Size = size; Byte = readByte; BeforeCreate = beforeCreate; RecordHistory = recordHistory;
            BeforeDiagnosticRecord = beforeDiagnosticRecord; BeforeDiagnosticSnapshot = beforeDiagnosticSnapshot;
        }
    }

    private readonly struct RawFontStyle
    {
        public int Os2Offset { get; }
        public int Os2Length { get; }
        public int HeadOffset { get; }
        public int HeadLength { get; }
        public ushort Selection { get; }
        public ushort MacStyle { get; }
        public bool Bold => (Selection & 0x20) != 0;
        public bool Italic => (Selection & 1) != 0;
        internal RawFontStyle(int os2Offset, int os2Length, int headOffset, int headLength, ushort selection, ushort macStyle)
        { Os2Offset = os2Offset; Os2Length = os2Length; HeadOffset = headOffset; HeadLength = headLength; Selection = selection; MacStyle = macStyle; }
    }

    // Only bounded structural/operation metadata. No resource ID, family name,
    // Unicode text or font payload is retained by the process-wide history.
    private sealed class FontReadStep
    {
        public long Generation { get; }
        public int Ordinal { get; }
        public int Thread { get; }
        public string Phase { get; }
        public string Operation { get; }
        public string? Tag { get; }
        public int? Offset { get; }
        public int? Size { get; }
        public bool? Success { get; }
        public byte? Value { get; }
        public int Wrapper { get; }
        public string Handle { get; }
        public long StartedTimestamp { get; }
        public long CompletedTimestamp { get; }
        public long TimestampFrequency => Stopwatch.Frequency;
        public int? OriginalLength { get; }
        public string? OriginalSha256 { get; }
        internal FontReadStep(long generation, int ordinal, string phase, string operation,
            uint? tag, int? offset, int? size, bool? success, byte? value, int wrapper, string handle,
            long? startedTimestamp, int? originalLength, string? originalSha256)
        {
            Generation = generation; Ordinal = ordinal; Thread = Environment.CurrentManagedThreadId;
            Phase = phase; Operation = operation; Tag = tag?.ToString("X8", CultureInfo.InvariantCulture);
            Offset = offset; Size = size; Success = success; Value = value; Wrapper = wrapper; Handle = handle;
            CompletedTimestamp = Stopwatch.GetTimestamp(); StartedTimestamp = startedTimestamp ?? CompletedTimestamp;
            OriginalLength = originalLength; OriginalSha256 = originalSha256;
        }
    }

    private sealed class FontReadFailureEvidence
    {
        public int SchemaVersion => 1;
        public string Classification { get; }
        public long Generation { get; }
        public bool Injected { get; }
        public bool DecoratedAfterOriginalThrow { get; }
        public bool RecordingFailed { get; }
        public int? OriginalLength { get; }
        public string? OriginalSha256 { get; }
        public bool RawStylePreflightPassed { get; }
        public object? RawStyle { get; }
        public object[] Steps { get; }
        public object[] ProcessHistory { get; }
        internal FontReadFailureEvidence(string classification, FontReadTrace trace, bool decoratedAfterOriginalThrow)
        {
            Classification = classification; Generation = trace.Generation; Injected = trace.Injected;
            DecoratedAfterOriginalThrow = decoratedAfterOriginalThrow;
            RecordingFailed = trace.RecordingFailed;
            OriginalLength = trace.Original?.Length;
            if (trace.Original is not null)
            {
                OriginalSha256 = trace.OriginalSha256 ?? FontReadTrace.Hash(trace.Original);
            }
            RawStylePreflightPassed = trace.Raw is not null; RawStyle = trace.Raw;
            Steps = trace._steps?.ToArray() ?? Array.Empty<object>();
            lock (FontReadTrace.HistoryGate) ProcessHistory = trace.HistoryEnabled ? FontReadTrace.History?.ToArray() ?? Array.Empty<object>() : Array.Empty<object>();
        }
    }

    private sealed class FontReadTrace
    {
        private static long _generation;
        private static readonly FontReadTrace Disabled = new();
        internal static readonly object HistoryGate = new();
        internal static Queue<object>? History;
        internal readonly long Generation;
        internal readonly byte[]? Original;
        internal readonly bool Injected;
        internal RawFontStyle? Raw;
        internal readonly List<object>? _steps;
        internal readonly bool HistoryEnabled;
        internal bool RecordingFailed;
        internal readonly string? OriginalSha256;
        private readonly bool _disabled;
        private readonly Action? _beforeRecord, _beforeSnapshot;
        private int _ordinal, _wrapper;
        private string _handle = "uncreated";
        private FontReadTrace() { _disabled = true; }
        private FontReadTrace(byte[]? original, NativeFontTableAccess? access)
        {
            Generation = Interlocked.Increment(ref _generation); Original = original; Injected = access is not null;
            _beforeRecord = access?.BeforeDiagnosticRecord; _beforeSnapshot = access?.BeforeDiagnosticSnapshot;
            _steps = new List<object>();
            HistoryEnabled = access?.RecordHistory ?? !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OFDRW_SKIA_DIAGNOSTICS_DIRECTORY"));
            // Optional cross-generation history needs content identity without
            // retaining arrays. This digest is diagnostic only; stream hashing
            // and byte comparison still run independently during admission.
            if (HistoryEnabled && original is not null)
            {
                try { OriginalSha256 = Hash(original); } catch { RecordingFailed = true; }
            }
            Record("create", "managed-clone");
        }
        internal static string Hash(byte[] original)
        { using var hash = SHA256.Create(); return BitConverter.ToString(hash.ComputeHash(original)).Replace("-", "").ToLowerInvariant(); }
        internal static FontReadTrace Create(byte[]? original, NativeFontTableAccess? access = null)
        { try { return new FontReadTrace(original, access); } catch { return Disabled; } }
        internal static long StartTimestamp()
        { try { return Stopwatch.GetTimestamp(); } catch { return 0; } }
        internal void SetRaw(RawFontStyle raw) { if (!_disabled) Raw = raw; }
        internal void Bind(SKTypeface face)
        {
            if (_disabled) return;
            try
            {
                _wrapper = RuntimeHelpers.GetHashCode(face);
                // Save while live; never query it to complete a failure report.
                _handle = face.Handle.ToInt64().ToString("X", CultureInfo.InvariantCulture);
                Record("create", "native-face");
            }
            catch { RecordingFailed = true; }
        }
        internal void Record(string phase, string operation, uint? tag = null, int? offset = null,
            int? size = null, bool? success = null, byte? value = null, long? startedTimestamp = null)
        {
            if (_disabled) return;
            try
            {
                _beforeRecord?.Invoke();
                var step = new FontReadStep(Generation, ++_ordinal, phase, operation, tag, offset, size, success, value,
                    _wrapper, _handle, startedTimestamp, Original?.Length, OriginalSha256);
                if (_steps!.Count == 32) _steps.RemoveAt(0);
                _steps.Add(step);
                if (HistoryEnabled) lock (HistoryGate)
                {
                    History ??= new Queue<object>();
                    if (History.Count == 128) History.Dequeue();
                    History.Enqueue(step);
                }
            }
            catch { RecordingFailed = true; }
        }
        internal void Attach(Exception error, string classification, bool decoratedAfterOriginalThrow = true)
        {
            if (_disabled) return;
            try
            {
                _beforeSnapshot?.Invoke();
                if (!error.Data.Contains(FontReadFailureKey))
                    error.Data[FontReadFailureKey] = new FontReadFailureEvidence(classification, this, decoratedAfterOriginalThrow);
            }
            catch { RecordingFailed = true; /* Keep the original failure. */ }
        }
        internal NotSupportedException Failure(string classification, string message)
        { var error = new NotSupportedException(message); Attach(error, classification, false); return error; }
    }

    private static int NativeTableSize(SKTypeface face, uint tag, string phase, FontReadTrace trace, NativeFontTableAccess? access)
    {
        int size;
        var started = FontReadTrace.StartTimestamp();
        try { size = access?.Size is { } read ? read(face, tag, phase) : face.GetTableSize(tag); }
        catch (Exception error)
        {
            trace.Record(phase, "size-threw", tag, startedTimestamp: started);
            trace.Attach(error, "native-table-call-threw");
            throw; // Ordinary propagation; no diagnostic retry or fresh throw.
        }
        trace.Record(phase, "size", tag, size: size, startedTimestamp: started);
        return size;
    }
}
