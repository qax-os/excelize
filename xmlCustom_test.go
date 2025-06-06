package excelize

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func Test_xlsxCustomProperty_getPropertyValue(t *testing.T) {
	type fields struct {
		FmtID    string
		PID      string
		Name     string
		Text     *TextValue
		Bool     *BoolValue
		Number   *NumberValue
		DateTime *FileTimeValue
	}
	tests := []struct {
		name   string
		fields fields
		want   interface{}
	}{
		{
			name: "TextValue",
			fields: fields{
				Text: &TextValue{
					Text: "text",
				},
			},
			want: "text",
		},
		{
			name: "BoolValue",
			fields: fields{
				Bool: &BoolValue{
					Bool: true,
				},
			},
			want: true,
		},
		{
			name: "NumberValue",
			fields: fields{
				Number: &NumberValue{
					Number: 1.0,
				},
			},
			want: 1.0,
		},
		{
			name: "FileTimeValue",
			fields: fields{
				DateTime: &FileTimeValue{
					DateTime: "2006-01-02T15:04:05Z",
				},
			},
			want: time.Date(2006, 1, 2, 15, 4, 5, 0, time.UTC),
		},
		{
			name: "InvalidFileTimeValue",
			fields: fields{
				DateTime: &FileTimeValue{
					DateTime: "invalid",
				},
			},
			want: nil,
		},
		{
			name: "NilValue",
			fields: fields{
				Text:     nil,
				Bool:     nil,
				Number:   nil,
				DateTime: nil,
			},
			want: nil,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := &xlsxCustomProperty{
				FmtID:    tt.fields.FmtID,
				PID:      tt.fields.PID,
				Name:     tt.fields.Name,
				Text:     tt.fields.Text,
				Bool:     tt.fields.Bool,
				Number:   tt.fields.Number,
				DateTime: tt.fields.DateTime,
			}
			assert.Equalf(t, tt.want, p.getPropertyValue(), "getPropertyValue()")
		})
	}
}
