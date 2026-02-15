using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Validation;
using Ofdrw.Net.EmrTechSpec.Models;

namespace Ofdrw.Net.EmrTechSpec.Abstractions;

public interface IEmrTechSpecValidator
{
    Task<ValidationReport> ValidateAsync(Stream ofdStream, EmrValidationProfile profile, CancellationToken cancellationToken = default);
}
