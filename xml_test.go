package excelize

import (
	"bytes"
	"fmt"
	"github.com/stretchr/testify/assert"
	"testing"
)

func Test_escapeText(t *testing.T) {
	tests := []struct {
		name    string
		str     string
		wantW   []byte
		wantErr assert.ErrorAssertionFunc
	}{
		{
			name:    "Double quote",
			str:     "\"",
			wantW:   escQuot,
			wantErr: assert.NoError,
		},
		{
			name:    "Single quote",
			str:     "'",
			wantW:   escApos,
			wantErr: assert.NoError,
		},
		{
			name:    "Ampersand",
			str:     "&",
			wantW:   escAmp,
			wantErr: assert.NoError,
		},
		{
			name:    "Less than",
			str:     "<",
			wantW:   escLT,
			wantErr: assert.NoError,
		},
		{
			name:    "More than",
			str:     ">",
			wantW:   escGT,
			wantErr: assert.NoError,
		},
		{
			name:    "Tab",
			str:     "\t",
			wantW:   escTab,
			wantErr: assert.NoError,
		},
		{
			name:    "Carriage return",
			str:     "\r",
			wantW:   escCR,
			wantErr: assert.NoError,
		},
		{
			name:    "Replacement character",
			str:     fmt.Sprintf("Broken%s", "\uFFFF"),
			wantW:   []byte("Broken\uFFFD"),
			wantErr: assert.NoError,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			w := &bytes.Buffer{}
			err := escapeText(w, []byte(tt.str))
			if !tt.wantErr(t, err, fmt.Sprintf("escapeText(%v, %s)", w, tt.str)) {
				return
			}
			assert.Equalf(t, string(tt.wantW), w.String(), "escapeText(%v, %s)", w, tt.str)
		})
	}
}

func TestIsInCharacterRange(t *testing.T) {
	tests := []struct {
		runeValue rune
		expected  bool
	}{
		{0xE000, true},    // Private Use Area (PUA) - Should return true
		{0xD800, false},   // Surrogate Pair Start - Should return false
		{0x20, true},      // Space - Should return true
		{0xFFFD, true},    // Last valid PUA character - Should return true
		{0x10FFFF, true},  // Highest valid Unicode character - Should return true
		{0x110000, false}, // Beyond valid Unicode range - Should return false
	}

	for _, test := range tests {
		result := isInCharacterRange(test.runeValue)
		if result != test.expected {
			t.Errorf("isInCharacterRange(%U) = %v; want %v", test.runeValue, result, test.expected)
		}
	}
}
