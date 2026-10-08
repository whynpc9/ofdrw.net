using System.Globalization;
using System.Text;
using SkiaSharp;

// Deterministic raw PDF producer for tests and real package consumption. Not product code.
internal sealed class PdfFixture : IDisposable
{
    private readonly byte[] _bytes; private readonly SKTypeface _face; private readonly SKFont _font;
    private readonly Dictionary<ushort, char> _characters = new();
    internal PdfFixture(byte[] font)
    { _bytes = font; using var data = SKData.CreateCopy(font); _face = SKTypeface.FromData(data); _font = new SKFont(_face, 1000); }
    internal string Hex(string text)
    {
        var output = new StringBuilder();
        foreach (var character in text)
        {
            var glyph = _face.GetGlyph(character);
            if (glyph == 0) throw new InvalidOperationException("Fixture font lacks " + character);
            _characters[glyph] = character; output.Append(glyph.ToString("X4"));
        }
        return output.ToString();
    }
    internal string Text(string text, double x, double y, double size = 12, double shear = 0)
        => $"BT /F1 {Number(size)} Tf 1 0 {Number(shear)} 1 {Number(x)} {Number(y)} Tm <{Hex(text)}> Tj ET\n";
    internal byte[] Create(string[] contents, string pageEntries = "", bool wrongUnicode = false, string fontEntries = "", bool inheritedResources = false)
    {
        var cmap = "/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def /CMapName /Fixture def /CMapType 2 def 1 begincodespacerange <0000> <FFFF> endcodespacerange " + _characters.Count + " beginbfchar " +
            string.Join(" ", _characters.OrderBy(pair => pair.Key).Select(pair => $"<{pair.Key:X4}> <{(int)(wrongUnicode && pair.Value == 'A' ? 'B' : pair.Value):X4}>")) + " endbfchar endcmap CMapName currentdict /CMap defineresource pop end end";
        var widths = string.Join(" ", _characters.OrderBy(pair => pair.Key).Select(pair => pair.Key + " [" + Number(_font.MeasureText(pair.Value.ToString())) + "]"));
        const string resources = "/Resources << /Font << /F1 3 0 R >> /XObject << /Im1 8 0 R >> /ExtGState << /GS1 << /ca .4 /BM /Multiply >> >> >>";
        var objects = new List<byte[]> {
            Ascii("<< /Type /Catalog /Pages 2 0 R >>"),
            Ascii("<< /Type /Pages /Count " + contents.Length + " /Kids [" + string.Join(" ", Enumerable.Range(0, contents.Length).Select(i => (9 + 2 * i) + " 0 R")) + "] " + (inheritedResources ? resources : "") + " >>"),
            Ascii("<< /Type /Font /Subtype /Type0 /BaseFont /FixtureFace /Encoding /Identity-H /DescendantFonts [4 0 R] /ToUnicode 6 0 R " + fontEntries + " >>"),
            Ascii("<< /Type /Font /Subtype /CIDFontType2 /BaseFont /FixtureFace /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 5 0 R /CIDToGIDMap /Identity /DW 1000 /W [" + widths + "] >>"),
            Ascii("<< /Type /FontDescriptor /FontName /FixtureFace /Flags 32 /FontBBox [-1000 -1000 3000 3000] /ItalicAngle 0 /Ascent 1000 /Descent -300 /CapHeight 700 /StemV 80 /FontFile2 7 0 R >>"),
            Stream(Ascii(cmap)), Stream(_bytes, "/Length1 " + _bytes.Length),
            Stream(new byte[] { 210, 40, 60, 30, 130, 190, 30, 130, 190, 210, 40, 60 }, "/Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8")
        };
        for (var index = 0; index < contents.Length; index++)
        {
            objects.Add(Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 420 595] " + (inheritedResources ? "" : resources) + " /Contents " + (10 + 2 * index) + " 0 R " + pageEntries + " >>"));
            objects.Add(Stream(Ascii(contents[index])));
        }
        using var output = new MemoryStream();
        void Write(string value) { var bytes = Ascii(value); output.Write(bytes); }
        Write("%PDF-1.7\n"); var offsets = new List<long>();
        for (var index = 0; index < objects.Count; index++)
        { offsets.Add(output.Position); Write((index + 1) + " 0 obj\n"); output.Write(objects[index]); Write("\nendobj\n"); }
        var xref = output.Position; Write("xref\n0 " + (objects.Count + 1) + "\n0000000000 65535 f \n");
        foreach (var offset in offsets) Write(offset.ToString("D10") + " 00000 n \n");
        Write("trailer\n<< /Size " + (objects.Count + 1) + " /Root 1 0 R >>\nstartxref\n" + xref + "\n%%EOF\n");
        return output.ToArray();
    }
    private static byte[] Stream(byte[] bytes, string extra = "")
    { using var output = new MemoryStream(); output.Write(Ascii("<< /Length " + bytes.Length + " " + extra + " >>\nstream\n")); output.Write(bytes); output.Write(Ascii("\nendstream")); return output.ToArray(); }
    private static byte[] Ascii(string value) => Encoding.ASCII.GetBytes(value);
    private static string Number(double value) => value.ToString("0.######", CultureInfo.InvariantCulture);
    public void Dispose() { _font.Dispose(); _face.Dispose(); }
}
