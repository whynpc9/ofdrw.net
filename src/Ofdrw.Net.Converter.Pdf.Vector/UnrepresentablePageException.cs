using System;

namespace Ofdrw.Net.Converter.Pdf.Vector;

// Rasterization cannot make an unrepresentable physical OFD page writable.
// Only the completed native-event path uses this classification.
internal sealed class UnrepresentablePageException : NotSupportedException
{
    internal UnrepresentablePageException(int page) : base("PDFV_UNSUPPORTED page=" + page +
        " PAGE_WRITER_PRECISION: physical page dimensions cannot remain positive at OFD writer precision; raster output has the same page-size requirement.") { }
}
