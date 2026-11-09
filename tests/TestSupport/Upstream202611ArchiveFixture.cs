using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace Ofdrw.Net.TestSupport;

internal static class Upstream202611ArchiveFixture
{
    private const string Namespace = "http://www.ofdspec.org/2016";
    private const string PagePath = "Doc_0/A/Content.xml";
    private const int PrefixSize = 77;

    // Solid one-pixel RGB PNGs. The same basename in both branches must not alias.
    internal static byte[] Red => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC");
    internal static byte[] Blue => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNgYPgPAAEDAQAIicLsAAAAAElFTkSuQmCC");

    internal static byte[] Create(string layout = "plain", string defect = "none")
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Entry(string path, byte[] data)
            {
                var entry = zip.CreateEntry(path, CompressionLevel.NoCompression);
                entry.LastWriteTime = new DateTimeOffset(2026, 11, 9, 0, 0, 0, TimeSpan.Zero);
                using var output = entry.Open();
                output.Write(data);
            }

            void Xml(string path, string element, string body, string attributes = "") =>
                Entry(path, Encoding.UTF8.GetBytes($"<{element} xmlns='{Namespace}'{attributes}>{body}</{element}>"));

            Xml("OFD.xml", "OFD", "<DocBody><DocInfo><DocID>plan019</DocID></DocInfo><DocRoot>dOC_0/document.xml</DocRoot></DocBody>",
                " Version='1.1' DocType='OFD'");
            Xml("Doc_0/Document.xml", "Document", """
                <CommonData>
                  <MaxUnitID>30</MaxUnitID>
                  <PageArea><PhysicalBox>0 0 100 70</PhysicalBox></PageArea>
                  <TemplatePage ID='3' BaseLoc='b/TPLS/template.xml' ZOrder='Background'/>
                </CommonData>
                <Pages><Page ID='1' BaseLoc='a/content.xml'/></Pages>
                """);
            if (defect != "missing-page")
                Xml(PagePath, "Page", "<Template TemplateID='3'/>" + Image(2, 4, 7, 10));
            Xml("DOC_0/B/Tpls/Template.xml", "Page", Image(5, 6, 8, 60));
            Xml("Doc_0/A/PageRes.xml", "Res", Media(7, "Swatch.png"), " BaseLoc='/doc_0/B'");
            Xml("DOC_0/B/Tpls/PageRes.xml", "Res", Media(8, "swatch.PNG"), " BaseLoc='/DOC_0/a'");
            Entry("Doc_0/A/swatch.png", Red);
            Entry("DOC_0/B/swatch.png", Blue);
            if (defect == "collision")
                Entry("DOC_0/a/CONTENT.XML", Encoding.UTF8.GetBytes("<wrong/>"));
            if (defect == "traversal")
                Entry("Doc_0/../../outside.xml", []);
        }

        var bytes = stream.ToArray();
        if (layout != "plain")
        {
            var prefix = Enumerable.Repeat((byte)0x5a, PrefixSize).ToArray();
            // A complete empty EOCD inside the preamble, followed by non-ZIP bytes.
            prefix.AsSpan(16, 22).Clear();
            "PK\u0005\u0006"u8.CopyTo(prefix.AsSpan(16));
            bytes = [.. prefix, .. bytes];
            if (layout == "adjusted")
                ShiftOffsets(bytes);
            else if (layout != "unadjusted")
                throw new ArgumentOutOfRangeException(nameof(layout));
        }

        return defect switch
        {
            "truncated" => bytes[..^8],
            "false-local-header" => CorruptLocalHeader(bytes),
            "none" or "missing-page" or "collision" or "traversal" => bytes,
            _ => throw new ArgumentOutOfRangeException(nameof(defect))
        };
    }

    private static string Image(int layer, int image, int resource, int x) =>
        $"<Content><Layer ID='{layer}'><ImageObject ID='{image}' ResourceID='{resource}' Boundary='{x} 20 30 30' CTM='30 0 0 30 0 0'/></Layer></Content>";

    private static string Media(int id, string name) =>
        $"<MultiMedias><MultiMedia ID='{id}' Type='Image' Format='PNG'><MediaFile>{name}</MediaFile></MultiMedia></MultiMedias>";

    private static void ShiftOffsets(byte[] bytes)
    {
        var footerOffsetField = bytes.AsSpan(bytes.Length - 6, 4);
        var centralOffset = BinaryPrimitives.ReadInt32LittleEndian(footerOffsetField) + PrefixSize;
        BinaryPrimitives.WriteInt32LittleEndian(footerOffsetField, centralOffset);
        for (var offset = centralOffset; offset < bytes.Length - 22;)
        {
            var localOffsetField = bytes.AsSpan(offset + 42, 4);
            BinaryPrimitives.WriteInt32LittleEndian(localOffsetField,
                BinaryPrimitives.ReadInt32LittleEndian(localOffsetField) + PrefixSize);
            offset += 46
                + BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 28, 2))
                + BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 30, 2))
                + BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 32, 2));
        }
    }

    private static byte[] CorruptLocalHeader(byte[] bytes)
    {
        // Keep the central directory intact, but replace the first local header signature
        // with a different PK record signature.
        bytes[2] = 0x05;
        bytes[3] = 0x06;
        return bytes;
    }
}
