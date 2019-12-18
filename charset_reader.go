package excelize

//import "gopkg.in/webnice/debug.v1"
//import "gopkg.in/webnice/log.v2"
import (
	"fmt"
	"io"
	"regexp"
	"strings"

	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
)

// CharsetReader Decoder from all codepages to UTF-8
func CharsetReader(charset string, input io.Reader) (rdr io.Reader, err error) {
	var (
		sm, nm                string
		item, enc             encoding.Encoding
		i                     int
		rexReplaceCharsetName = regexp.MustCompile(`[-_ ]`)
	)

	sm = rexReplaceCharsetName.ReplaceAllString(strings.ToLower(charset), ``)
	for i = range charmap.All {
		item, nm = charmap.All[i], rexReplaceCharsetName.ReplaceAllString(strings.ToLower(fmt.Sprintf("%s", item)), ``)
		if strings.EqualFold(sm, nm) {
			enc = item
		}
	}
	if enc == nil {
		err = fmt.Errorf("CharsetReader: unexpected charset: %q", charset)
		return
	}
	rdr = enc.NewDecoder().Reader(input)

	return
}
