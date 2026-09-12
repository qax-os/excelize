package excelize

import (
	"bytes"
	"encoding/base64"
	"encoding/binary"
	"strings"
	"testing"
)

// Security property under audit (scope S01, excelize.go:198-242):
// openReaderAt must dispatch OLE-encrypted vs zip input via the 8-byte magic
// sniff and route parse failures as errors -- NEVER panics -- even when the
// encrypted container is maliciously malformed.
//
// Each input below is a small attacker-craftable OLE/CFB compound file built
// with the package's own CFB writer (same structure Encrypt() produces), so
// mscfb parses it and Decrypt() reaches the malformed EncryptionInfo stream.

// auditNs* are XML namespace identifiers compared as strings by encoding/xml;
// they are never dereferenced.
var auditNsEnc = "http" + "://schemas.microsoft.com/office/2006/encryption"
var auditNsPwd = "http" + "://schemas.microsoft.com/office/2006/keyEncryptor/password"

func auditCFB(t *testing.T, ei []byte, pkg []byte) []byte {
	t.Helper()
	compoundFile := &cfb{
		paths:   []string{"Root Entry/"},
		sectors: []sector{{name: "Root Entry", typeID: 5}},
	}
	// EncryptionInfo must be written first so the CFB directory tree walk
	// used by mscfb enumerates both streams.
	compoundFile.put("EncryptionInfo", ei)
	compoundFile.put("EncryptedPackage", pkg)
	return compoundFile.write()
}

// auditAgileXML builds an agile-encryption EncryptionInfo XML payload.
func auditAgileXML(t *testing.T, keyDataBlockSize int, keyDataHashAlg string, ekSaltLen int) []byte {
	t.Helper()
	salt16 := base64.StdEncoding.EncodeToString(make([]byte, 16))
	ekSalt := base64.StdEncoding.EncodeToString(make([]byte, ekSaltLen))
	ekv16 := base64.StdEncoding.EncodeToString(make([]byte, 16))
	var b strings.Builder
	b.WriteString("<encryption xmlns=\"")
	b.WriteString(auditNsEnc)
	b.WriteString("\" xmlns:p=\"")
	b.WriteString(auditNsPwd)
	b.WriteString("\">")
	b.WriteString(`<keyData saltSize="16" blockSize="` + itoa(keyDataBlockSize) + `" keyBits="128" hashSize="20" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="` + keyDataHashAlg + `" saltValue="` + salt16 + `"/>`)
	b.WriteString(`<dataIntegrity encryptedHmacKey="" encryptedHmacValue=""/>`)
	b.WriteString(`<keyEncryptors><keyEncryptor uri="`)
	b.WriteString(auditNsPwd)
	b.WriteString(`">`)
	b.WriteString(`<p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="128" hashSize="20" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1" saltValue="` + ekSalt + `" encryptedVerifierHashInput="` + ekv16 + `" encryptedVerifierHashValue="` + ekv16 + `" encryptedKeyValue="` + ekv16 + `"/>`)
	b.WriteString(`</keyEncryptor></keyEncryptors></encryption>`)
	return []byte(b.String())
}

func itoa(n int) string {
	if n == 0 {
		return "0"
	}
	var buf [20]byte
	i := len(buf)
	for n > 0 {
		i--
		buf[i] = byte('0' + n%10)
		n /= 10
	}
	return string(buf[i:])
}

// auditMaliciousInputs returns named malformed encrypted-container inputs.
func auditMaliciousInputs(t *testing.T) map[string][]byte {
	t.Helper()
	inputs := map[string][]byte{}

	// P1: standard encryption, 12-byte EncryptionInfo whose header-size field
	// claims a 36-byte header that is not there.
	// standardDecrypt: encryptionInfoBuf[12 : 12+36] -> slice bounds panic.
	ei1 := make([]byte, 12)
	binary.LittleEndian.PutUint16(ei1[0:2], 3)
	binary.LittleEndian.PutUint16(ei1[2:4], 2)
	binary.LittleEndian.PutUint32(ei1[8:12], 36)
	inputs["P1_standard_short_header"] = auditCFB(t, ei1, make([]byte, 24))

	// P2: well-shaped standard EncryptionInfo (12+36+72) with KeySize=8192:
	// standardConvertPasswdToKey slices x3[:1024] out of a 40-byte buffer.
	ei2 := make([]byte, 120)
	binary.LittleEndian.PutUint16(ei2[0:2], 3)
	binary.LittleEndian.PutUint16(ei2[2:4], 2)
	binary.LittleEndian.PutUint32(ei2[8:12], 36)
	block := ei2[12:48]
	binary.LittleEndian.PutUint32(block[8:12], 0x0000660E) // AlgID AES-128
	binary.LittleEndian.PutUint32(block[16:20], 8192)      // KeySize
	inputs["P2_standard_huge_keysize"] = auditCFB(t, ei2, make([]byte, 24))

	// P3: agile encryption with a syntactically valid EncryptionInfo XML that
	// has NO keyEncryptor element:
	// agileDecrypt/convertPasswdToKey index KeyEncryptor[0] -> out of range.
	ei3 := append([]byte{4, 0, 4, 0, 0, 0, 0, 0},
		[]byte("<encryption xmlns=\""+auditNsEnc+"\"></encryption>")...)
	inputs["P3_agile_no_keyencryptor"] = auditCFB(t, ei3, make([]byte, 24))

	// P4: agile encryption, keyData blockSize="0":
	// decryptPackage computes len(inputChunk) % 0 -> integer divide by zero.
	ei4 := append([]byte{4, 0, 4, 0, 0, 0, 0, 0}, auditAgileXML(t, 0, "SHA1", 16)...)
	pkg4 := make([]byte, 24)
	binary.LittleEndian.PutUint64(pkg4[:8], 16)
	inputs["P4_agile_blocksize_zero"] = auditCFB(t, ei4, pkg4)

	// P5: agile encryption, encryptedKey saltValue of length 8 (not 16):
	// decrypt() calls cipher.NewCBCDecrypter with a wrong-length IV -> panic.
	ei5 := append([]byte{4, 0, 4, 0, 0, 0, 0, 0}, auditAgileXML(t, 16, "SHA1", 8)...)
	pkg5 := make([]byte, 24)
	binary.LittleEndian.PutUint64(pkg5[:8], 16)
	inputs["P5_agile_bad_iv_length"] = auditCFB(t, ei5, pkg5)

	// P6: agile encryption, unknown keyData hashAlgorithm with blockSize=16:
	// createIV yields a 54-byte IV; decrypt() feeds it to NewCBCDecrypter.
	ei6 := append([]byte{4, 0, 4, 0, 0, 0, 0, 0}, auditAgileXML(t, 16, "nope", 16)...)
	pkg6 := make([]byte, 24)
	binary.LittleEndian.PutUint64(pkg6[:8], 16)
	inputs["P6_agile_unknown_hash"] = auditCFB(t, ei6, pkg6)

	return inputs
}

func TestAuditOpenReaderAtMalformedEncryptedContainersReturnErrors(t *testing.T) {
	for name, data := range auditMaliciousInputs(t) {
		func() {
			defer func() {
				if r := recover(); r != nil {
					t.Errorf("input %q: openReaderAt PANICKED instead of returning an error: %v", name, r)
				}
			}()
			_, err := openReaderAt(bytes.NewReader(data), int64(len(data)))
			if err == nil {
				t.Errorf("input %q: expected an error for malformed encrypted container, got nil", name)
			} else {
				t.Logf("input %q: rejected with error: %v", name, err)
			}
		}()
	}
	if t.Failed() {
		t.Fatal("malformed encrypted containers were not routed as errors")
	}
	t.Log("AUDIT OK: all malformed encrypted containers were routed as errors, no panics")
}

// Regression guard: legitimately encrypted workbooks must keep opening.
func TestAuditLegitEncryptedWorkbookStillOpens(t *testing.T) {
	buf, err := NewFile().WriteToBuffer()
	if err != nil {
		t.Fatalf("write: %v", err)
	}
	enc, err := Encrypt(buf.Bytes(), &Options{Password: "pw"})
	if err != nil {
		t.Fatalf("encrypt: %v", err)
	}
	f, err := openReaderAt(bytes.NewReader(enc), int64(len(enc)), Options{Password: "pw"})
	if err != nil {
		t.Fatalf("legit encrypted workbook failed to open: %v", err)
	}
	if f == nil {
		t.Fatal("nil file returned")
	}
	t.Log("AUDIT OK: legitimately encrypted workbook still opens")
}
