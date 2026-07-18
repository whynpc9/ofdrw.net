using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Signatures.Crypto;

namespace Ofdrw.Net.Signatures.Verification;

public sealed class OfdSignatureVerifier
{
    private readonly Dictionary<string, IOfdSignedValueVerifier> _signedValueVerifiers;

    public OfdSignatureVerifier(
        IEnumerable<IOfdSignedValueVerifier>? signedValueVerifiers = null)
    {
        _signedValueVerifiers = (signedValueVerifiers ??
                Array.Empty<IOfdSignedValueVerifier>())
            .ToDictionary(
                verifier => verifier.SignatureMethod,
                StringComparer.OrdinalIgnoreCase);
    }

    public async Task<OfdSignatureVerificationReport> VerifyAsync(
        Stream ofdInput,
        CancellationToken cancellationToken = default)
    {
        if (ofdInput is null)
        {
            throw new ArgumentNullException(nameof(ofdInput));
        }

        var archive = await new OfdPackageLoader()
            .LoadAsync(ofdInput, cancellationToken)
            .ConfigureAwait(false);
        var report = new OfdSignatureVerificationReport();
        if (!archive.Contains("OFD.xml"))
        {
            report.Issues.Add("OFD.xml is missing.");
            return report;
        }

        XDocument ofd;
        try
        {
            ofd = XDocument.Parse(
                archive.ReadUtf8Text("OFD.xml"),
                LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception)
        {
            report.Issues.Add($"OFD.xml could not be parsed: {exception.Message}");
            return report;
        }

        var signatureLocations = ofd
            .Descendants()
            .Where(element => element.Name.LocalName == "Signatures")
            .Select(element => element.Value)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (signatureLocations.Count == 0)
        {
            report.Issues.Add("The document does not declare a signatures list.");
            return report;
        }

        foreach (var location in signatureLocations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await VerifySignatureListAsync(
                    archive,
                    location,
                    report,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        return report;
    }

    private async Task VerifySignatureListAsync(
        OfdPackageArchive archive,
        string location,
        OfdSignatureVerificationReport report,
        CancellationToken cancellationToken)
    {
        var listPath = NormalizePath(location);
        if (!archive.TryGetBytes(listPath, out var listBytes))
        {
            report.Issues.Add($"Signatures list is missing: {location}");
            return;
        }

        XDocument list;
        try
        {
            list = XDocument.Parse(
                System.Text.Encoding.UTF8.GetString(listBytes).TrimStart('\uFEFF'),
                LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception)
        {
            report.Issues.Add(
                $"Signatures list '{listPath}' could not be parsed: {exception.Message}");
            return;
        }

        foreach (var record in list
            .Descendants()
            .Where(element => element.Name.LocalName == "Signature"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var baseLocation = record.Attribute("BaseLoc")?.Value;
            var result = new OfdSignatureVerificationResult
            {
                Id = record.Attribute("ID")?.Value ?? string.Empty
            };
            report.Signatures.Add(result);
            if (string.IsNullOrWhiteSpace(baseLocation))
            {
                result.CryptographicStatus = OfdCryptographicVerificationStatus.Error;
                result.CryptographicMessage = "Signature record has no BaseLoc.";
                continue;
            }

            var signaturePath = ResolvePath(listPath, baseLocation!);
            result.SignaturePath = signaturePath;
            if (!archive.TryGetBytes(signaturePath, out var signatureBytes))
            {
                result.CryptographicStatus = OfdCryptographicVerificationStatus.Missing;
                result.CryptographicMessage =
                    $"Signature description is missing: {signaturePath}";
                continue;
            }

            await VerifySignatureAsync(
                    archive,
                    signatureBytes,
                    signaturePath,
                    result,
                    cancellationToken)
                .ConfigureAwait(false);
        }
    }

    private async Task VerifySignatureAsync(
        OfdPackageArchive archive,
        byte[] signatureBytes,
        string signaturePath,
        OfdSignatureVerificationResult result,
        CancellationToken cancellationToken)
    {
        XDocument signature;
        try
        {
            signature = XDocument.Parse(
                System.Text.Encoding.UTF8.GetString(signatureBytes).TrimStart('\uFEFF'),
                LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception)
        {
            result.CryptographicStatus = OfdCryptographicVerificationStatus.Error;
            result.CryptographicMessage =
                $"Signature description could not be parsed: {exception.Message}";
            return;
        }

        result.SignatureMethod = signature
            .Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "SignatureMethod")
            ?.Value ?? string.Empty;
        var references = signature
            .Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "References");
        result.CheckMethod = references?.Attribute("CheckMethod")?.Value ?? string.Empty;
        foreach (var reference in references?
                     .Elements()
                     .Where(element => element.Name.LocalName == "Reference") ??
                 Enumerable.Empty<XElement>())
        {
            var fileReference = reference.Attribute("FileRef")?.Value ?? string.Empty;
            var referenceResult = new OfdReferenceVerificationResult
            {
                FileReference = fileReference,
                ResolvedEntryName = NormalizePath(fileReference),
                DigestAlgorithmSupported =
                    OfdDigestAlgorithms.IsSupported(result.CheckMethod)
            };
            result.References.Add(referenceResult);
            if (!TryNormalizePath(fileReference, out var entryName))
            {
                referenceResult.Error = "The reference path is unsafe.";
                continue;
            }

            referenceResult.ResolvedEntryName = entryName;
            if (!archive.TryGetBytes(entryName, out var entryBytes))
            {
                referenceResult.Error = "The referenced package entry is missing.";
                continue;
            }

            referenceResult.EntryExists = true;
            if (!referenceResult.DigestAlgorithmSupported)
            {
                referenceResult.Error =
                    $"Unsupported digest algorithm: {result.CheckMethod}";
                continue;
            }

            try
            {
                var expected = Convert.FromBase64String(
                    reference.Elements()
                        .FirstOrDefault(element =>
                            element.Name.LocalName == "CheckValue")
                        ?.Value ?? string.Empty);
                var actual = OfdDigestAlgorithms.Compute(
                    result.CheckMethod,
                    entryBytes);
                referenceResult.DigestMatches = FixedTimeEquals(expected, actual);
            }
            catch (Exception exception)
            {
                referenceResult.Error = exception.Message;
            }
        }

        var signedValueLocation = signature
            .Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "SignedValue")
            ?.Value;
        if (string.IsNullOrWhiteSpace(signedValueLocation))
        {
            result.CryptographicStatus = OfdCryptographicVerificationStatus.Missing;
            result.CryptographicMessage = "SignedValue is not declared.";
            return;
        }

        var signedValuePath = ResolvePath(signaturePath, signedValueLocation!);
        result.SignedValuePath = signedValuePath;
        if (!archive.TryGetBytes(signedValuePath, out var signedValue))
        {
            result.CryptographicStatus = OfdCryptographicVerificationStatus.Missing;
            result.CryptographicMessage =
                $"Signed value is missing: {signedValuePath}";
            return;
        }

        if (!_signedValueVerifiers.TryGetValue(
                result.SignatureMethod,
                out var signedValueVerifier))
        {
            result.CryptographicStatus =
                OfdCryptographicVerificationStatus.Unsupported;
            result.CryptographicMessage =
                $"No signed-value verifier is registered for {result.SignatureMethod}. " +
                "Reference digests were checked independently.";
            return;
        }

        try
        {
            var propertyInformation = "/" + signaturePath;
            var valid = await signedValueVerifier
                .VerifyAsync(
                    signatureBytes,
                    signedValue,
                    propertyInformation,
                    cancellationToken)
                .ConfigureAwait(false);
            result.CryptographicStatus = valid
                ? OfdCryptographicVerificationStatus.Valid
                : OfdCryptographicVerificationStatus.Invalid;
            result.CryptographicMessage = valid
                ? "The signed value is valid."
                : "The signed value is invalid.";
        }
        catch (Exception exception)
        {
            result.CryptographicStatus = OfdCryptographicVerificationStatus.Error;
            result.CryptographicMessage = exception.Message;
        }
    }

    private static bool FixedTimeEquals(byte[] left, byte[] right)
    {
        var difference = left.Length ^ right.Length;
        var length = Math.Min(left.Length, right.Length);
        for (var index = 0; index < length; index++)
        {
            difference |= left[index] ^ right[index];
        }

        return difference == 0;
    }

    private static string ResolvePath(string containingPath, string reference)
    {
        if (reference.StartsWith("/", StringComparison.Ordinal))
        {
            return NormalizePath(reference);
        }

        var slash = containingPath.LastIndexOf('/');
        var directory = slash < 0 ? string.Empty : containingPath.Substring(0, slash + 1);
        return NormalizePath(directory + reference);
    }

    private static string NormalizePath(string path)
    {
        return path.Replace('\\', '/').TrimStart('/');
    }

    private static bool TryNormalizePath(string path, out string normalized)
    {
        normalized = NormalizePath(path);
        return !normalized.Split('/').Any(segment => segment == "..");
    }
}
