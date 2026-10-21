package main

import (
    "bytes"
    "crypto/sha256"
    "fmt"
    "os"

    "github.com/tjfoc/gmsm/sm3"
    "github.com/tjfoc/gmsm/x509"
)

func main() {
    if len(os.Args) != 5 { panic("expected .NET Signature.xml, .NET SignedData, Java Signature.xml, Java SignedData") }
    for i := 1; i < len(os.Args); i += 2 {
        xml, err := os.ReadFile(os.Args[i]); if err != nil { panic(err) }
        name := os.Args[i+1]
        data, err := os.ReadFile(name); if err != nil { panic(err) }
        p7, err := x509.ParsePKCS7(data); if err != nil { panic(err) }
        if len(p7.Signers) != 1 || len(p7.Certificates) != 1 { panic(fmt.Errorf("%s: expected one signer/certificate", name)) }
        // This library's PKCS7 Verify() accepts OID .501, while the pinned Java
        // GB/T 35275 container uses .301.1. Parse independently and export the
        // exact content, certificate and signature for OpenSSL SM2 verification.
        signer := p7.Signers[0]
        if signer.DigestEncryptionAlgorithm.Algorithm.String() != "1.2.156.10197.1.301.1" {
            panic(fmt.Errorf("%s: unexpected signer algorithm", name))
        }
        if !bytes.Equal(p7.Content, sm3.Sm3Sum(xml)) { panic(fmt.Errorf("%s: attached SM3 digest does not match XML", name)) }
        for suffix, payload := range map[string][]byte{".content": p7.Content, ".cert": p7.Certificates[0].Raw, ".sig": signer.EncryptedDigest} {
            if err := os.WriteFile(name+suffix, payload, 0600); err != nil { panic(err) }
        }
        fmt.Printf("%s: independent SignedData parse and attached SM3 digest verified, content=%x, certs=%d, bytes=%d, xml-sha256=%x\n", name, p7.Content, len(p7.Certificates), len(data), sha256.Sum256(xml))
        if bytes.Equal(p7.Content, xml) { panic("attached content is raw XML, expected digest") }
    }
}
