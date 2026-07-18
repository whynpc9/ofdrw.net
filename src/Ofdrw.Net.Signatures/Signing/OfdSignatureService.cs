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

        options ??= new OfdSignatureOptions();
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

        var signature = BuildSignatureXml(
            entries,
            signaturesPath,
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
    }

    private static XDocument BuildSignatureXml(
        IReadOnlyDictionary<string, byte[]> entries,
        string signaturesPath,
        string signaturePath,
        string signedValuePath,
        IOfdSignatureProvider provider,
        OfdSignatureOptions options)
    {
        var ns = XNamespace.Get("http://www.ofdspec.org/2016");
        var references = new XElement(
            ns + "References",
            new XAttribute("CheckMethod", options.CheckMethod));
        foreach (var entry in entries
            .Where(pair =>
                options.ProtectSignatureList ||
                !string.Equals(
                    pair.Key,
                    signaturesPath,
                    StringComparison.OrdinalIgnoreCase))
            .OrderBy(pair => pair.Key, StringComparer.Ordinal))
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
            new XElement(ns + "SignatureMethod", provider.SignatureMethod),
            new XElement(
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

        return signaturesDirectory.TrimEnd('/') + "/Sign_" + index;
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
