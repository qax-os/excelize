package excelize

type Row struct {
	w *Worksheet
	row *xlsxRow
}

func (r *Row)SetHeight(height float64) {
	r.row.Ht = height
	r.row.CustomHeight = true
}

func (r *Row)SetVisible(visible bool) {
	if visible {
		r.row.Hidden = false
		return
	}

	r.row.Hidden = true
}

func (r *Row) GetVisible() bool {
	return !r.row.Hidden
}
