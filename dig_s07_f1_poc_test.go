// dig_s07_f1_poc_test.go — regression test for the extractPart allocation
// bound in crypt.go: a version-4 CFB whose "EncryptionInfo" directory entry
// declares streamSize 0xFFFFFFFFFFFFFFFF makes mscfb report entry.Size == -1,
// which extractPart previously fed straight into make([]byte, entry.Size)
// (panic: makeslice: len out of range, escaping Decrypt).
package excelize

import (
	"encoding/binary"
	"fmt"
	"testing"
	"unicode/utf16"
)

// recoverPanic runs body and reports whether a panic escaped it.
func recoverPanic(body func()) (panicked bool, msg string) {
	defer func() {
		if r := recover(); r != nil {
			panicked = true
			msg = fmt.Sprint(r)
		}
	}()
	body()
	return
}

// buildV4CFBNegativeStreamSize builds a version-4 (4096-byte sector) compound
// file whose "EncryptionInfo" directory entry declares streamSize
// 0xFFFFFFFFFFFFFFFF and owns no data sectors (startingSector = endOfChain).
func buildV4CFBNegativeStreamSize() []byte {
	const (
		endOfChain = uint32(0xFFFFFFFE)
		noStream   = uint32(0xFFFFFFFF)
		fatSectID  = uint32(0xFFFFFFFD)
	)
	hdr := make([]byte, 512)
	binary.LittleEndian.PutUint64(hdr[0:8], 0xE11AB1A1E011CFD0) // signature
	binary.LittleEndian.PutUint16(hdr[24:26], 0x003E)           // minor version
	binary.LittleEndian.PutUint16(hdr[26:28], 0x0004)           // major version 4
	binary.LittleEndian.PutUint16(hdr[30:32], 0x000C)           // sector shift -> 4096
	binary.LittleEndian.PutUint16(hdr[32:34], 0x0006)           // mini sector shift
	binary.LittleEndian.PutUint32(hdr[40:44], 1)                // num directory sectors
	binary.LittleEndian.PutUint32(hdr[44:48], 1)                // num FAT sectors
	binary.LittleEndian.PutUint32(hdr[48:52], 0)                // directory sector = 0
	binary.LittleEndian.PutUint32(hdr[56:60], 4096)             // mini stream cutoff
	binary.LittleEndian.PutUint32(hdr[60:64], endOfChain)       // mini FAT start
	binary.LittleEndian.PutUint32(hdr[64:68], 0)                // num mini FAT sectors
	binary.LittleEndian.PutUint32(hdr[68:72], endOfChain)       // DIFAT start
	binary.LittleEndian.PutUint32(hdr[72:76], 0)                // num DIFAT sectors
	binary.LittleEndian.PutUint32(hdr[76:80], 1)                // DIFAT[0] = sector 1 (FAT)

	dir := make([]byte, 4096)
	putEntry := func(off int, name string, objType byte, left, right, child, start uint32, size uint64) {
		e := dir[off : off+128]
		encoded := utf16.Encode([]rune(name + "\x00"))
		for i, r := range encoded {
			binary.LittleEndian.PutUint16(e[i*2:i*2+2], r)
		}
		binary.LittleEndian.PutUint16(e[64:66], uint16((len(name)+1)*2))
		e[66] = objType
		e[67] = 1 // black
		binary.LittleEndian.PutUint32(e[68:72], left)
		binary.LittleEndian.PutUint32(e[72:76], right)
		binary.LittleEndian.PutUint32(e[76:80], child)
		binary.LittleEndian.PutUint32(e[116:120], start)
		binary.LittleEndian.PutUint64(e[120:128], size)
	}
	putEntry(0, "Root Entry", 5, noStream, noStream, 1, endOfChain, 0)
	putEntry(128, "EncryptionInfo", 2, noStream, noStream, noStream, endOfChain, 0xFFFFFFFFFFFFFFFF)

	fat := make([]byte, 4096)
	binary.LittleEndian.PutUint32(fat[0:4], endOfChain) // directory chain ends
	binary.LittleEndian.PutUint32(fat[4:8], fatSectID)  // FAT sector itself

	out := append([]byte{}, hdr...)
	for len(out) < 4096 {
		out = append(out, 0)
	}
	out = append(out, dir...)
	out = append(out, fat...)
	return out
}

// TestDigS07F1ExtractPartAllocation drives the public Decrypt API with the
// malformed container above. Before the fix: "F1 BUG REPRODUCED: panic
// escaped Decrypt(): runtime error: makeslice: len out of range".
// After the fix: "F1 BLOCKED: Decrypt rejected malformed CFB with error".
func TestDigS07F1ExtractPartAllocation(t *testing.T) {
	raw := buildV4CFBNegativeStreamSize()
	panicked, msg := recoverPanic(func() {
		packageBuf, err := Decrypt(raw, &Options{Password: "password"})
		if err != nil {
			fmt.Printf("F1 BLOCKED: Decrypt rejected malformed CFB with error: %v\n", err)
			return
		}
		fmt.Printf("F1 BLOCKED: no panic (%d bytes)\n", len(packageBuf))
	})
	if panicked {
		fmt.Printf("F1 BUG REPRODUCED: panic escaped Decrypt(): %v\n", msg)
	}
}
