using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.Validation;
using Ofdrw.Net.EmrTechSpec.Services;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.EmrTechSpec.Tests;

public sealed class EmrValidationTests
{
    [Fact]
    public async Task Validator_ShouldPassCoreRules_ForValidGeneratedOfd()
    {
        var builder = new OfdDocumentBuilder();
        builder.AddPage(new Ofdrw.Net.Core.Models.OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement
                {
                    Text = "EMR Sample",
                    FontName = "SimSun",
                    FontSizeMillimeters = 4,
                    XMillimeters = 10,
                    YMillimeters = 12,
                    WidthMillimeters = 80,
                    HeightMillimeters = 10
                }
            }
        });

        var writer = new OfdPackageWriter();
        await using var ms = new MemoryStream();
        await writer.WriteAsync(builder.Build(), ms);

        var profileRepo = new EmrValidationProfileRepository();
        var profile = profileRepo.GetDefaultProfile();
        var validator = new EmrTechSpecValidator();

        ms.Position = 0;
        var report = await validator.ValidateAsync(ms, profile);

        Assert.True(report.Passed);
        Assert.DoesNotContain(report.Findings, x => x.Severity == ValidationSeverity.Error);
    }

    [Fact]
    public async Task Validator_ShouldDetectWrongDocType()
    {
        var builder = new OfdDocumentBuilder();
        builder.SetOptions(new OfdDocumentOptions { DocType = "OFD-X" });
        builder.AddPage(new Ofdrw.Net.Core.Models.OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297
        });

        var writer = new OfdPackageWriter();
        await using var ms = new MemoryStream();
        await writer.WriteAsync(builder.Build(), ms);

        var profileRepo = new EmrValidationProfileRepository();
        var profile = profileRepo.GetDefaultProfile();
        var validator = new EmrTechSpecValidator();

        ms.Position = 0;
        var report = await validator.ValidateAsync(ms, profile);

        Assert.Contains(report.Findings, x => x.RuleId == "r-doc-type-ofd-h" && x.Severity == ValidationSeverity.Error);
        Assert.False(report.Passed);
    }
}
