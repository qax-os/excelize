package excelize

//import "gopkg.in/webnice/log.v2"
import (
	"bytes"
	"io"
	"io/ioutil"
	"testing"

	"golang.org/x/text/encoding/charmap"
)

func TestCharsetReaderSupported(t *testing.T) {
	const unknownCodepage = `DEE26011`
	var (
		err                error
		inp                io.Reader
		out                io.Reader
		i                  int
		supportedCodepages = []string{
			charmap.CodePage037.String(),
			charmap.CodePage437.String(),
			charmap.CodePage037.String(),
			charmap.CodePage850.String(),
			charmap.CodePage852.String(),
			charmap.CodePage855.String(),
			charmap.CodePage858.String(),
			charmap.CodePage860.String(),
			charmap.CodePage862.String(),
			charmap.CodePage863.String(),
			charmap.CodePage865.String(),
			charmap.CodePage866.String(),
			charmap.CodePage1047.String(),
			charmap.CodePage1140.String(),
			charmap.ISO8859_1.String(),
			charmap.ISO8859_2.String(),
			charmap.ISO8859_3.String(),
			charmap.ISO8859_4.String(),
			charmap.ISO8859_5.String(),
			charmap.ISO8859_6.String(),
			charmap.ISO8859_7.String(),
			charmap.ISO8859_8.String(),
			charmap.ISO8859_9.String(),
			charmap.ISO8859_10.String(),
			charmap.ISO8859_13.String(),
			charmap.ISO8859_14.String(),
			charmap.ISO8859_15.String(),
			charmap.ISO8859_16.String(),
			charmap.KOI8R.String(),
			charmap.KOI8U.String(),
			charmap.Macintosh.String(),
			charmap.MacintoshCyrillic.String(),
			charmap.Windows874.String(),
			charmap.Windows1250.String(),
			charmap.Windows1251.String(),
			charmap.Windows1252.String(),
			charmap.Windows1253.String(),
			charmap.Windows1254.String(),
			charmap.Windows1255.String(),
			charmap.Windows1256.String(),
			charmap.Windows1257.String(),
			charmap.Windows1258.String(),
		}
	)

	inp = bytes.NewReader([]byte{})
	// Unknown codepage
	if out, err = CharsetReader(unknownCodepage, inp); err == nil || out != nil {
		t.Fatalf("incorrect CharsetReader implementation")
	}
	// Known codepages
	for i = range supportedCodepages {
		if out, err = CharsetReader(supportedCodepages[i], inp); err != nil || out == nil {
			t.Fatalf("codepage %q is not supported", supportedCodepages[i])
		}
	}
}

func TestCharsetReaderTranslation(t *testing.T) {
	var (
		err         error
		out         io.Reader
		buf         []byte
		destination = []byte{
			0xD0, 0x9F, 0xD1, 0x80, 0xD0, 0xB8, 0xD0, 0xB2, 0xD0, 0xB5,
			0xD1, 0x82, 0x20, 0xD0, 0xBC, 0xD0, 0xB8, 0xD1, 0x80,
		}
		tests = []struct {
			Source      []byte
			Destination []byte
			Name        string
		}{
			{[]byte{0xF0, 0xD2, 0xC9, 0xD7, 0xC5, 0xD4, 0x20, 0xCD, 0xC9, 0xD2}, destination, "koi8-r"},
			{[]byte{0x8F, 0xE0, 0xA8, 0xA2, 0xA5, 0xE2, 0x20, 0xAC, 0xA8, 0xE0}, destination, "ibm code page 866"},
			{[]byte{0xCF, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2, 0x20, 0xEC, 0xE8, 0xF0}, destination, "windows-1251"},
			{[]byte{0x8F, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2, 0x20, 0xEC, 0xE8, 0xF0}, destination, "macintosh cyrillic"},
		}
	)

	for _, test := range tests {
		if out, err = CharsetReader(test.Name, bytes.NewReader(test.Source)); err != nil {
			t.Fatalf("incorrect CharsetReader implementation, error: %s", err)
		}
		if buf, err = ioutil.ReadAll(out); err != nil {
			t.Fatalf("ioutil ReadAll error: %s", err)
		}
		if !bytes.EqualFold(buf, test.Destination) {
			t.Fatalf("transcode error, incorrect decoding %s", test.Name)
		}
		buf, out = buf[:0], nil
	}
}
