using System;
using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdFontSelection
{
    internal static OfdFontResource? Resolve(IEnumerable<OfdFontResource> fonts, OfdTextElement text)
    {
        var values = fonts as IReadOnlyList<OfdFontResource> ?? fonts.ToArray();
        var explicitFont = values.FirstOrDefault(font => !string.IsNullOrEmpty(text.FontResourceId) && font.Id == text.FontResourceId);
        var explicitId = OfdSourceIdentity.ParseId(text.FontResourceId);
        if (explicitFont is null && explicitId > 0)
            explicitFont = values.FirstOrDefault(font => OfdSourceIdentity.ParseId(font.Id) == explicitId);
        if (explicitFont is not null) return explicitFont;
        var family = values.Where(font => string.Equals(font.FontName, text.FontName, StringComparison.OrdinalIgnoreCase));
        return family.FirstOrDefault(font => font.Bold == (text.Weight >= 600) && font.Italic == text.Italic)
            ?? family.FirstOrDefault(font => !font.Bold && !font.Italic) ?? family.FirstOrDefault();
    }
}
