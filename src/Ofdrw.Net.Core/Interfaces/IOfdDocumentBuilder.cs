using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Interfaces;

public interface IOfdDocumentBuilder
{
    OfdDocumentOptions Options { get; }

    IOfdDocumentBuilder SetOptions(OfdDocumentOptions options);

    IOfdDocumentBuilder AddPage(OfdPage page);

    IOfdDocumentBuilder AddText(int pageIndex, OfdTextElement element);

    IOfdDocumentBuilder AddImage(int pageIndex, OfdImageElement element);

    IOfdDocumentBuilder AddAttachment(OfdAttachment attachment);

    OfdDocumentPackage Build();
}
