using System;
using System.IO;
using System.Linq;
using System.Text.Json.Serialization;
using System.Text.Json;
using System.Reflection;
using Ofdrw.Net.EmrTechSpec.Models;

namespace Ofdrw.Net.EmrTechSpec.Services;

public sealed class EmrValidationProfileRepository
{
    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters =
        {
            new JsonStringEnumConverter()
        }
    };

    public EmrValidationProfile GetDefaultProfile()
    {
        var assembly = typeof(EmrValidationProfileRepository).GetTypeInfo().Assembly;
        var embeddedResource = assembly.GetManifestResourceNames()
            .FirstOrDefault(x => x.EndsWith("Profiles.emr-ofd-h-202x.json", StringComparison.OrdinalIgnoreCase));
        if (embeddedResource is not null)
        {
            using var stream = assembly.GetManifestResourceStream(embeddedResource);
            if (stream is not null)
            {
                using var reader = new StreamReader(stream);
                var embeddedJson = reader.ReadToEnd();
                var embeddedProfile = JsonSerializer.Deserialize<EmrValidationProfile>(embeddedJson, _jsonOptions);
                if (embeddedProfile is not null)
                {
                    return embeddedProfile;
                }
            }
        }

        var basePath = AppContext.BaseDirectory;
        var profilePath = Path.Combine(basePath, "Profiles", "emr-ofd-h-202x.json");
        if (!File.Exists(profilePath))
        {
            profilePath = Path.GetFullPath(Path.Combine(basePath, "..", "..", "..", "..", "src", "Ofdrw.Net.EmrTechSpec", "Profiles", "emr-ofd-h-202x.json"));
        }

        if (!File.Exists(profilePath))
        {
            throw new FileNotFoundException("Profile file not found.", profilePath);
        }

        var json = File.ReadAllText(profilePath);
        var profile = JsonSerializer.Deserialize<EmrValidationProfile>(json, _jsonOptions);
        if (profile is null)
        {
            throw new InvalidDataException("Invalid profile content.");
        }

        return profile;
    }
}
