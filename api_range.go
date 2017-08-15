package excelize

import (
	"encoding/json"

	"github.com/plandem/excelize/format"
)

type Range struct {
	w *Worksheet
	fromAxis string
	toAxis string
}

func (r *Range)walk(cb func(c *xlsxC)) {
	for _, row := range r.w.sheet.SheetData.Row {
		for _, c := range row.C {
			if checkCellInArea(c.R, r.fromAxis + ":" + r.toAxis) {
				cb(c)
			}
		}
	}
}

func (r *Range)Clear() {
	r.walk(func(c *xlsxC) { c.V = "" })
}

func (r *Range)Contains(axis string) bool {
	return checkCellInArea(axis, r.fromAxis + ":" + r.toAxis)
}

func (r *Range)SetStyle(s Style) {
	r.walk(func(c *xlsxC) { c.S = int(s) })
}

func (r *Range)SetConditionalFormat(formatSet string){
	var fs []*format.Conditional
	json.Unmarshal([]byte(formatSet), &fs)

	drawContFmtFunc := map[string]func(p int, ct string, fmtCond *format.Conditional) *xlsxCfRule{
		"cellIs":          drawCondFmtCellIs,
		"top10":           drawCondFmtTop10,
		"aboveAverage":    drawCondFmtAboveAverage,
		"duplicateValues": drawCondFmtDuplicateUniqueValues,
		"uniqueValues":    drawCondFmtDuplicateUniqueValues,
		"2_color_scale":   drawCondFmtColorScale,
		"3_color_scale":   drawCondFmtColorScale,
		"dataBar":         drawCondFmtDataBar,
	}

	cfRule := []*xlsxCfRule{}
	for p, v := range fs {
		var vt, ct string
		var ok bool
		// "type" is a required parameter, check for valid validation types.
		vt, ok = validType[v.Type]
		if !ok {
			continue
		}
		// Check for valid criteria types.
		ct, ok = criteriaType[v.Criteria]
		if !ok {
			continue
		}

		drawfunc, ok := drawContFmtFunc[vt]
		if ok {
			cfRule = append(cfRule, drawfunc(p, ct, v))
		}
	}

	r.w.sheet.ConditionalFormatting = append(r.w.sheet.ConditionalFormatting, &xlsxConditionalFormatting{
		SQRef:  r.fromAxis + ":" + r.toAxis,
		CfRule: cfRule,
	})
}

