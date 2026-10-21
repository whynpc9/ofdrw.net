using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Signatures.Crypto;

namespace Ofdrw.Net.Signatures.Signing;

public sealed class OfdSignatureService
{
    public async Task SignAsync(
        Stream ofdInput,
        Stream ofdOutput,
        IOfdSignatureProvider provider,
        OfdSignatureOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        await SignWithResultAsync(ofdInput, ofdOutput, provider, options, cancellationToken)
            .ConfigureAwait(false);
    }

    /// <summary>Signs an OFD and returns the exact protected and unprotected final package entries.</summary>
    public async Task<OfdSignatureCreationResult> SignWithResultAsync(
        Stream ofdInput,
        Stream ofdOutput,
        IOfdSignatureProvider provider,
        OfdSignatureOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (ofdInput is null)
        {
            throw new ArgumentNullException(nameof(ofdInput));
        }

        if (ofdOutput is null)
        {
            throw new ArgumentNullException(nameof(ofdOutput));
        }

        if (provider is null)
        {
            throw new ArgumentNullException(nameof(provider));
        }

        options = SnapshotOptions(options ?? new OfdSignatureOptions());
        if (!OfdDigestAlgorithms.IsSupported(options.CheckMethod))
        {
            throw new NotSupportedException(
                $"Unsupported OFD digest algorithm: {options.CheckMethod}");
        }

        var archive = await new OfdPackageLoader()
            .LoadAsync(ofdInput, cancellationToken)
            .ConfigureAwait(false);
        var entries = archive.EntryNames.ToDictionary(
            name => name,
            name => archive.GetBytes(name),
            StringComparer.OrdinalIgnoreCase);
        if (!entries.TryGetValue("OFD.xml", out var ofdBytes))
        {
            throw new InvalidDataException("OFD.xml is missing.");
        }

        var ofd = ParseXml(ofdBytes, "OFD.xml");
        ValidateExtensions(options, ofd);
        var documentBody = ofd
            .Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "DocBody")
            ?? throw new InvalidDataException("OFD.xml has no DocBody.");
        var signaturesElement = documentBody
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName == "Signatures");
        string signaturesPath;
        if (signaturesElement is null)
        {
            var documentRoot = documentBody
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "DocRoot")
                ?.Value;
            if (string.IsNullOrWhiteSpace(documentRoot))
            {
                throw new InvalidDataException("OFD.xml has no DocRoot.");
            }

            signaturesPath =
                GetDirectory(NormalizePath(documentRoot!)) + "Signs/Signatures.xml";
            signaturesElement = new XElement(
                documentBody.Name.Namespace + "Signatures",
                "/" + signaturesPath);
            documentBody.Add(signaturesElement);
            entries["OFD.xml"] = SerializeXml(ofd);
        }
        else
        {
            signaturesPath = NormalizePath(signaturesElement.Value);
        }

        XDocument signatures;
        if (entries.TryGetValue(signaturesPath, out var existingSignaturesBytes))
        {
            signatures = ParseXml(existingSignaturesBytes, signaturesPath);
            EnsureListCanBeUpdated(entries, signatures, signaturesPath);
        }
        else
        {
            signatures = new XDocument(
                new XDeclaration("1.0", "utf-8", null),
                new XElement(
                    documentBody.Name.Namespace + "Signatures",
                    new XElement(
                        documentBody.Name.Namespace + "MaxSignId",
                        "0")));
        }

        var listRoot = signatures.Root ??
            throw new InvalidDataException("Signatures.xml has no root element.");
        var signatureRecords = listRoot
            .Elements()
            .Where(element => element.Name.LocalName == "Signature")
            .ToList();
        var nextId = NextNumericId(signatureRecords);
        var signaturesDirectory = GetDirectory(signaturesPath);
        var signDirectory = NextSignDirectory(
            entries.Keys,
            signaturesDirectory);
        var signaturePath = signDirectory + "/Signature.xml";
        var signedValuePath = signDirectory + "/SignedValue.dat";

        var record = new XElement(
            listRoot.Name.Namespace + "Signature",
            new XAttribute("ID", nextId),
            new XAttribute("Type", provider.SignatureType),
            new XAttribute(
                "BaseLoc",
                signaturePath.Substring(signaturesDirectory.Length)));
        if (options.RelativeSignatureId is not null)
        {
            var relativeId = options.RelativeSignatureId;
            var matchingBases = signatureRecords.Where(item => string.Equals(
                item.Attribute("ID")?.Value, relativeId, StringComparison.Ordinal)).ToArray();
            if (string.IsNullOrWhiteSpace(relativeId) || matchingBases.Length != 1)
                throw new InvalidDataException("Relative must identify exactly one existing signature in this list.");
            EnsureRelativeBaseExists(entries, matchingBases[0], signaturesPath);
            record.SetAttributeValue("Relative", relativeId);
        }
        listRoot.Add(record);
        var maxSignId = listRoot
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName == "MaxSignId");
        if (maxSignId is null)
        {
            listRoot.AddFirst(
                new XElement(listRoot.Name.Namespace + "MaxSignId", nextId));
        }
        else
        {
            maxSignId.Value = nextId.ToString();
        }

        entries[signaturesPath] = SerializeXml(signatures);
        if (provider.SealData is { Length: > 0 })
        {
            entries[signDirectory + "/Seal.esl"] =
                (byte[])provider.SealData.Clone();
        }

        var selection = OfdSignatureFileSelection.Select(
            entries, ofd, signaturesPath, options, cancellationToken);
        var protectedEntries = selection.Protected.ToDictionary(
            path => path, path => entries[path], StringComparer.OrdinalIgnoreCase);

        var signature = BuildSignatureXml(
            protectedEntries,
            signaturePath,
            signedValuePath,
            provider,
            options);
        var signatureBytes = SerializeXml(signature);
        var propertyInformation = "/" + signaturePath;
        var signedValue = await provider
            .SignAsync(signatureBytes, propertyInformation, cancellationToken)
            .ConfigureAwait(false);
        if (signedValue is null || signedValue.Length == 0)
        {
            throw new InvalidDataException(
                "The signature provider returned an empty signed value.");
        }

        entries[signaturePath] = signatureBytes;
        entries[signedValuePath] = signedValue;
        await WritePackageAsync(entries, ofdOutput, cancellationToken)
            .ConfigureAwait(false);
        return new OfdSignatureCreationResult(
            selection.Protected.Select(path => "/" + path).ToArray(),
            selection.Unprotected.Concat(new[] { signaturePath, signedValuePath })
                .Select(path => "/" + path).OrderBy(path => path, StringComparer.Ordinal).ToArray());
    }

    private static XDocument BuildSignatureXml(
        IReadOnlyDictionary<string, byte[]> entries,
        string signaturePath,
        string signedValuePath,
        IOfdSignatureProvider provider,
        OfdSignatureOptions options)
    {
        var ns = XNamespace.Get("http://www.ofdspec.org/2016");
        var references = new XElement(
            ns + "References",
            new XAttribute("CheckMethod", options.CheckMethod));
        foreach (var entry in entries.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            references.Add(
                new XElement(
                    ns + "Reference",
                    new XAttribute("FileRef", "/" + entry.Key),
                    new XElement(
                        ns + "CheckValue",
                        Convert.ToBase64String(
                            OfdDigestAlgorithms.Compute(
                                options.CheckMethod,
                                entry.Value)))));
        }

        var signedInfo = new XElement(
            ns + "SignedInfo",
            new XElement(
                ns + "Provider",
                new XAttribute("ProviderName", provider.ProviderName),
                new XAttribute("Company", provider.Company),
                new XAttribute("Version", provider.Version)),
            new XElement(ns + "SignatureMethod", provider.SignatureMethod));
        if (options.Parameters.Count > 0)
        {
            signedInfo.Add(new XElement(ns + "Parameters",
                options.Parameters.Select(parameter =>
                    new XElement(ns + "Parameter",
                        new XAttribute("Name", parameter.Name),
                        parameter.Type is null ? null : new XAttribute("Type", parameter.Type),
                        parameter.Value))));
        }
        signedInfo.Add(new XElement(
            ns + "SignatureDateTime",
            DateTime.UtcNow.ToString(
                "yyyy-MM-ddTHH:mm:ss'Z'",
                System.Globalization.CultureInfo.InvariantCulture)));
        if (provider.SealData is { Length: > 0 })
        {
            signedInfo.Add(
                new XElement(
                    ns + "Seal",
                    new XAttribute(
                        "BaseLoc",
                        "/" + GetDirectory(signaturePath) + "Seal.esl")));
        }

        signedInfo.Add(references);
        return new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            new XElement(
                ns + "Signature",
                signedInfo,
                new XElement(ns + "SignedValue", "/" + signedValuePath)));
    }

    private static void ValidateExtensions(OfdSignatureOptions options, XDocument ofd)
    {
        if (options.RelativeSignatureId is null && options.Parameters.Count == 0) return;
        if (!string.Equals(ofd.Root?.Attribute("Version")?.Value, "2.0", StringComparison.Ordinal))
            throw new NotSupportedException("Relative and Parameters require an OFD 2.0 document.");
        if (ofd.Root?.Name.NamespaceName != "http://www.ofdspec.org/2016")
            throw new NotSupportedException("OFD 2.0 signature extensions require the 2016 namespace.");
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var parameter in options.Parameters)
        {
            if (parameter is null || string.IsNullOrWhiteSpace(parameter.Name) ||
                parameter.Name.Length > 256 || parameter.Value.Length > 4096 ||
                parameter.Type?.Length > 256 || !names.Add(parameter.Name))
                throw new InvalidDataException("Invalid or duplicate signature Parameter name/value/type.");
        }
    }

    private static OfdSignatureOptions SnapshotOptions(OfdSignatureOptions source)
    {
        var copy = new OfdSignatureOptions
        {
            CheckMethod = source.CheckMethod,
            ProtectSignatureList = source.ProtectSignatureList,
            ProtectFile = source.ProtectFile,
            RelativeSignatureId = source.RelativeSignatureId
        };
        foreach (var parameter in source.Parameters) copy.Parameters.Add(parameter);
        return copy;
    }

    private static void EnsureRelativeBaseExists(
        IReadOnlyDictionary<string, byte[]> entries, XElement record, string signaturesPath)
    {
        var baseLocation = record.Attribute("BaseLoc")?.Value;
        if (string.IsNullOrWhiteSpace(baseLocation))
            throw new InvalidDataException("Relative base has no Signature.xml location.");
        var basePath = baseLocation!.StartsWith("/", StringComparison.Ordinal)
            ? NormalizePath(baseLocation)
            : GetDirectory(signaturesPath) + NormalizePath(baseLocation);
        if (!entries.TryGetValue(basePath, out var bytes) || bytes.Length == 0)
            throw new InvalidDataException("Relative base Signature.xml is missing.");
        var baseSignature = ParseXml(bytes, basePath);
        var valueLocation = baseSignature.Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "SignedValue")?.Value;
        if (string.IsNullOrWhiteSpace(valueLocation))
            throw new InvalidDataException("Relative base has no SignedValue location.");
        var valuePath = valueLocation!.StartsWith("/", StringComparison.Ordinal)
            ? NormalizePath(valueLocation)
            : GetDirectory(basePath) + NormalizePath(valueLocation);
        if (!entries.TryGetValue(valuePath, out var value) || value.Length == 0)
            throw new InvalidDataException("Relative base SignedValue is missing.");
    }

    private static void EnsureListCanBeUpdated(
        IReadOnlyDictionary<string, byte[]> entries,
        XDocument signatures,
        string signaturesPath)
    {
        foreach (var record in signatures
            .Descendants()
            .Where(element => element.Name.LocalName == "Signature"))
        {
            var baseLocation = record.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(baseLocation))
            {
                continue;
            }

            var signaturePath = baseLocation!.StartsWith("/", StringComparison.Ordinal)
                ? NormalizePath(baseLocation)
                : GetDirectory(signaturesPath) + NormalizePath(baseLocation);
            if (!entries.TryGetValue(signaturePath, out var signatureBytes))
            {
                continue;
            }

            var signature = ParseXml(signatureBytes, signaturePath);
            var protectsList = signature
                .Descendants()
                .Where(element => element.Name.LocalName == "Reference")
                .Select(element => NormalizePath(
                    element.Attribute("FileRef")?.Value ?? string.Empty))
                .Any(path => string.Equals(
                    path,
                    signaturesPath,
                    StringComparison.OrdinalIgnoreCase));
            if (protectsList)
            {
                throw new InvalidOperationException(
                    $"The existing signature '{signaturePath}' protects " +
                    $"{signaturesPath}; appending would invalidate it.");
            }
        }
    }

    private static int NextNumericId(IReadOnlyCollection<XElement> records)
    {
        var max = 0;
        foreach (var value in records.Select(record => record.Attribute("ID")?.Value))
        {
            if (int.TryParse(value, out var parsed))
            {
                max = Math.Max(max, parsed);
            }
        }

        return checked(max + 1);
    }

    private static string NextSignDirectory(
        IEnumerable<string> entryNames,
        string signaturesDirectory)
    {
        var index = 0;
        while (entryNames.Any(name => name.StartsWith(
            signaturesDirectory + "Sign_" + index + "/",
            StringComparison.OrdinalIgnoreCase)))
        {
            index++;
        }

        // GetDirectory already includes the separator for a non-root list.
        return signaturesDirectory + "Sign_" + index;
    }

    private static async Task WritePackageAsync(
        IReadOnlyDictionary<string, byte[]> entries,
        Stream output,
        CancellationToken cancellationToken)
    {
        using var archive = new ZipArchive(
            output,
            ZipArchiveMode.Create,
            leaveOpen: true);
        foreach (var entry in entries.OrderBy(
            pair => pair.Key,
            StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var zipEntry = archive.CreateEntry(
                entry.Key,
                CompressionLevel.Optimal);
            using var stream = zipEntry.Open();
            await stream
                .WriteAsync(
                    entry.Value,
                    0,
                    entry.Value.Length,
                    cancellationToken)
                .ConfigureAwait(false);
        }
    }

    private static XDocument ParseXml(byte[] bytes, string path)
    {
        try
        {
            return XDocument.Parse(
                Encoding.UTF8.GetString(bytes).TrimStart('\uFEFF'),
                LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception)
        {
            throw new InvalidDataException(
                $"Could not parse {path}.",
                exception);
        }
    }

    private static byte[] SerializeXml(XDocument document)
    {
        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(
            stream,
            new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = false,
                OmitXmlDeclaration = false
            }))
        {
            document.Save(writer);
        }

        return stream.ToArray();
    }

    private static string GetDirectory(string path)
    {
        var normalized = NormalizePath(path);
        var index = normalized.LastIndexOf('/');
        return index < 0 ? string.Empty : normalized.Substring(0, index + 1);
    }

    private static string NormalizePath(string path)
    {
        var normalized = path.Replace('\\', '/').TrimStart('/');
        if (normalized.Split('/').Any(segment => segment == ".."))
        {
            throw new InvalidDataException($"Unsafe OFD path: {path}");
        }

        return normalized;
    }
}
