package workbook

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Workbook is the in-memory model used by the engine. Rows keep display values
// for searching and previews, while TypedValues keeps values that should be
// written back as numbers, booleans, nil, time values, etc.
type Workbook struct {
	Sheets         []Sheet
	SourcePath     string
	OriginalSheets []Sheet
	TypedValues    map[string]map[string]any
}

type Sheet struct {
	Name string
	Rows [][]string
}

type FindResult struct {
	Sheet    string   `json:"sheet"`
	Address  string   `json:"address"`
	RowIndex int      `json:"rowIndex"`
	ColIndex int      `json:"colIndex"`
	Value    string   `json:"value"`
	RowData  []string `json:"rowData"`
}

type CellChange struct {
	Sheet    string `json:"sheet"`
	Cell     string `json:"cell"`
	RowIndex int    `json:"rowIndex"`
	ColIndex int    `json:"colIndex"`
	OldValue string `json:"oldValue"`
	NewValue string `json:"newValue"`
}

type Diff struct {
	ChangedCells int          `json:"changedCells"`
	ChangedRows  int          `json:"changedRows"`
	Changes      []CellChange `json:"changes"`
}

func (d *Diff) AddCell(sheet string, rowIdx, colIdx int, oldValue, newValue string) {
	cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
	if err != nil {
		cell = fmt.Sprintf("R%dC%d", rowIdx+1, colIdx+1)
	}
	d.ChangedCells++
	d.Changes = append(d.Changes, CellChange{
		Sheet:    sheet,
		Cell:     cell,
		RowIndex: rowIdx + 1,
		ColIndex: colIdx + 1,
		OldValue: oldValue,
		NewValue: newValue,
	})
}

func New() *Workbook {
	return &Workbook{}
}

func FromSheets(sheets []Sheet) *Workbook {
	return &Workbook{Sheets: CloneSheets(sheets)}
}

func (b *Workbook) SheetByName(name string) *Sheet {
	for i := range b.Sheets {
		if b.Sheets[i].Name == name {
			return &b.Sheets[i]
		}
	}
	return nil
}

func (b *Workbook) SetSource(path string) {
	b.SourcePath = NormalizePath(path)
	b.OriginalSheets = CloneSheets(b.Sheets)
	b.TypedValues = nil
}

func (b *Workbook) MarkSaved(path string) {
	b.SetSource(path)
}

func (b *Workbook) RememberCellValue(sheet string, rowIdx, colIdx int, value any) {
	if rowIdx < 0 || colIdx < 0 {
		return
	}
	cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
	if err != nil {
		return
	}
	b.RememberCellValueByName(sheet, cell, value)
}

func (b *Workbook) RememberCellValueByName(sheet, cell string, value any) {
	if b.TypedValues == nil {
		b.TypedValues = make(map[string]map[string]any)
	}
	if b.TypedValues[sheet] == nil {
		b.TypedValues[sheet] = make(map[string]any)
	}
	b.TypedValues[sheet][strings.ToUpper(cell)] = value
}

func (b *Workbook) RememberedCellValue(sheet string, rowIdx, colIdx int) (any, bool) {
	if b.TypedValues == nil || rowIdx < 0 || colIdx < 0 {
		return nil, false
	}
	cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
	if err != nil {
		return nil, false
	}
	values, ok := b.TypedValues[sheet]
	if !ok {
		return nil, false
	}
	value, ok := values[strings.ToUpper(cell)]
	return value, ok
}

func (b *Workbook) ClearSheetTypedValues(sheet string) {
	if b.TypedValues != nil {
		delete(b.TypedValues, sheet)
	}
}

func (b *Workbook) ClearRowTypedValues(sheet string, rowIdx int) {
	if b.TypedValues == nil || rowIdx < 0 {
		return
	}
	values, ok := b.TypedValues[sheet]
	if !ok {
		return
	}
	rowNum := rowIdx + 1
	for cell := range values {
		_, row, err := excelize.CellNameToCoordinates(cell)
		if err == nil && row == rowNum {
			delete(values, cell)
		}
	}
}

func (s *Sheet) EnsureSize(rowIdx, colIdx int) {
	for len(s.Rows) <= rowIdx {
		s.Rows = append(s.Rows, []string{})
	}
	if len(s.Rows[rowIdx]) <= colIdx {
		newRow := make([]string, colIdx+1)
		copy(newRow, s.Rows[rowIdx])
		s.Rows[rowIdx] = newRow
	}
}

func (s *Sheet) Cell(rowIdx, colIdx int) string {
	if rowIdx < 0 || rowIdx >= len(s.Rows) {
		return ""
	}
	if colIdx < 0 || colIdx >= len(s.Rows[rowIdx]) {
		return ""
	}
	return s.Rows[rowIdx][colIdx]
}

func (s *Sheet) SetCell(rowIdx, colIdx int, value string) {
	s.EnsureSize(rowIdx, colIdx)
	s.Rows[rowIdx][colIdx] = value
}

func CloneRows(rows [][]string) [][]string {
	copied := make([][]string, len(rows))
	for i, row := range rows {
		copied[i] = append([]string(nil), row...)
	}
	return copied
}

func CloneSheets(sheets []Sheet) []Sheet {
	copied := make([]Sheet, len(sheets))
	for i, sheet := range sheets {
		copied[i] = Sheet{Name: sheet.Name, Rows: CloneRows(sheet.Rows)}
	}
	return copied
}

func NormalizePath(path string) string {
	abs, err := filepath.Abs(path)
	if err != nil {
		return path
	}
	return abs
}

func SamePath(left, right string) bool {
	return strings.EqualFold(NormalizePath(left), NormalizePath(right))
}

func DisplayValue(value any) string {
	if value == nil {
		return ""
	}
	return fmt.Sprintf("%v", value)
}

func MaxColumnCount(rows [][]string) int {
	maxCols := 0
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}
	return maxCols
}

func SheetRowsCount(sheets []Sheet) int {
	total := 0
	for _, sheet := range sheets {
		total += len(sheet.Rows)
	}
	return total
}
