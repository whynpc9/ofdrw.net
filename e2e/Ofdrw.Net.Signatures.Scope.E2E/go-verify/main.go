package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/base64"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
)

type signatureRecord struct {
	ID       string `xml:"ID,attr"`
	Relative string `xml:"Relative,attr"`
}
type signatureList struct {
	Records []signatureRecord `xml:"Signature"`
}
type parameter struct {
	Name  string `xml:"Name,attr"`
	Type  string `xml:"Type,attr"`
	Value string `xml:",chardata"`
}
type reference struct {
	Path   string `xml:"FileRef,attr"`
	Digest string `xml:"CheckValue"`
}
type signatureDocument struct {
	SignedInfo struct {
		Parameters struct {
			Items []parameter `xml:"Parameter"`
		} `xml:"Parameters"`
		References struct {
			Method string      `xml:"CheckMethod,attr"`
			Items  []reference `xml:"Reference"`
		} `xml:"References"`
	} `xml:"SignedInfo"`
}

func entries(path string) (map[string][]byte, error) {
	archive, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer archive.Close()
	result := make(map[string][]byte)
	for _, file := range archive.File {
		if _, exists := result[file.Name]; exists {
			return nil, fmt.Errorf("duplicate entry: %s", file.Name)
		}
		stream, err := file.Open()
		if err != nil {
			return nil, err
		}
		data, err := io.ReadAll(stream)
		stream.Close()
		if err != nil {
			return nil, err
		}
		result[file.Name] = data
	}
	return result, nil
}

func main() {
	if len(os.Args) != 4 {
		panic("expected first OFD, final OFD, output JSON")
	}
	first, err := entries(os.Args[1])
	if err != nil {
		panic(err)
	}
	last, err := entries(os.Args[2])
	if err != nil {
		panic(err)
	}
	for _, path := range []string{"Doc_0/Signs/Sign_0/Signature.xml", "Doc_0/Signs/Sign_0/SignedValue.dat"} {
		if !bytes.Equal(first[path], last[path]) {
			panic("old signature bytes changed: " + path)
		}
	}
	var list signatureList
	if err := xml.Unmarshal(last["Doc_0/Signs/Signatures.xml"], &list); err != nil {
		panic(err)
	}
	if len(list.Records) != 2 || list.Records[1].Relative != list.Records[0].ID {
		panic("Relative does not identify the earlier signature")
	}
	var current signatureDocument
	if err := xml.Unmarshal(last["Doc_0/Signs/Sign_1/Signature.xml"], &current); err != nil {
		panic(err)
	}
	properties := current.SignedInfo.Parameters.Items
	if len(properties) != 1 || properties[0].Name != "purpose" ||
		properties[0].Type != "text" || properties[0].Value != "package-consumer" {
		panic("SignedInfo/Parameters did not round-trip")
	}
	refs := current.SignedInfo.References
	if refs.Method != "2.16.840.1.101.3.4.2.1" {
		panic("expected SHA-256 references")
	}
	seen := make(map[string]bool)
	for _, ref := range refs.Items {
		if !strings.HasPrefix(ref.Path, "/") || strings.Contains(ref.Path, "..") {
			panic("unsafe FileRef")
		}
		path := strings.TrimPrefix(ref.Path, "/")
		if seen[path] {
			panic("duplicate FileRef")
		}
		seen[path] = true
		data, ok := last[path]
		if !ok {
			panic("missing referenced file: " + path)
		}
		expected, err := base64.StdEncoding.DecodeString(strings.TrimSpace(ref.Digest))
		if err != nil {
			panic(err)
		}
		actual := sha256.Sum256(data)
		if !bytes.Equal(expected, actual[:]) {
			panic("digest mismatch: " + path)
		}
	}
	if !seen["OFD.xml"] || !seen["Doc_0/Document.xml"] || seen["Extras/optional.bin"] {
		panic("incorrect required/optional reference selection")
	}
	summary := map[string]any{"independentGoReferencesChecked": len(refs.Items),
		"relativeMatchesBaseID": true, "parametersRoundTrip": true,
		"oldSignatureBytesUnchanged": true, "optionalEntryUnprotected": true}
	encoded, _ := json.MarshalIndent(summary, "", "  ")
	if err := os.WriteFile(os.Args[3], append(encoded, '\n'), 0644); err != nil {
		panic(err)
	}
	fmt.Println("Independent Go XML and SHA-256 checks passed.")
}
