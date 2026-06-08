package excelagent

import (
	"context"
	"fmt"

	"github.com/iEvan-lhr/go-excel-agent/workbook"
	"github.com/xuri/excelize/v2"
)

// Expose CellView and CellStyle as public type aliases for excelagent package consumers
type CellView = workbook.CellView
type CellStyle = workbook.CellStyle

// GetSheetView retrieves a 2D grid of CellView objects containing data, styles, and sizes for the specified sheet.
func (b *Book) GetSheetView(ctx context.Context, sheetName string) ([][]CellView, error) {
	if err := b.ensureEngine(); err != nil {
		return nil, err
	}
	sheet := b.engine.Book.SheetByName(sheetName)
	if sheet == nil {
		return nil, fmt.Errorf("找不到 sheet: %s", sheetName)
	}

	// If there's no source path, fall back to default sheet view
	if b.engine.Book.SourcePath == "" {
		return b.getDefaultSheetView(sheet), nil
	}

	file, err := excelize.OpenFile(b.engine.Book.SourcePath)
	if err != nil {
		return b.getDefaultSheetView(sheet), nil
	}
	defer file.Close()

	// Read rows from the excelize file directly
	rows, err := file.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("读取 sheet '%s' 失败: %w", sheetName, err)
	}

	// Fetch merge cell information
	type mergeInfo struct {
		rowSpan    int
		colSpan    int
		isOverlaid bool
		isMerged   bool
	}
	mergeMap := make(map[string]*mergeInfo)
	mergeCells, err := file.GetMergeCells(sheetName)
	maxMergeRow := 0
	maxMergeCol := 0
	if err == nil {
		for i := range mergeCells {
			mc := &mergeCells[i]
			startCol, startRow, err1 := excelize.CellNameToCoordinates(mc.GetStartAxis())
			endCol, endRow, err2 := excelize.CellNameToCoordinates(mc.GetEndAxis())
			if err1 == nil && err2 == nil {
				if endRow > maxMergeRow {
					maxMergeRow = endRow
				}
				if endCol > maxMergeCol {
					maxMergeCol = endCol
				}
				// Top-left cell of the merge
				topLeftCell, _ := excelize.CoordinatesToCellName(startCol, startRow)
				mergeMap[topLeftCell] = &mergeInfo{
					rowSpan:  endRow - startRow + 1,
					colSpan:  endCol - startCol + 1,
					isMerged: true,
				}
				// Overlaid cells that are covered by this merge range
				for r := startRow; r <= endRow; r++ {
					for c := startCol; c <= endCol; c++ {
						if r == startRow && c == startCol {
							continue
						}
						cellName, _ := excelize.CoordinatesToCellName(c, r)
						mergeMap[cellName] = &mergeInfo{
							isOverlaid: true,
							isMerged:   true,
						}
					}
				}
			}
		}
	}

	// Determine the maximum rows and columns
	maxRows := len(sheet.Rows)
	if len(rows) > maxRows {
		maxRows = len(rows)
	}
	if maxMergeRow > maxRows {
		maxRows = maxMergeRow
	}

	maxCols := 0
	for _, row := range sheet.Rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}
	if maxMergeCol > maxCols {
		maxCols = maxMergeCol
	}

	view := make([][]CellView, maxRows)
	for rowIdx := 0; rowIdx < maxRows; rowIdx++ {
		view[rowIdx] = make([]CellView, maxCols)

		// Get row height (default is 15.0)
		height, _ := file.GetRowHeight(sheetName, rowIdx+1)
		if height == 0 {
			height = 15.0
		}

		for colIdx := 0; colIdx < maxCols; colIdx++ {
			cellName, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
			if err != nil {
				continue
			}

			colName, err := excelize.ColumnNumberToName(colIdx + 1)
			if err != nil {
				continue
			}

			// Get column width (default is 8.43)
			width, _ := file.GetColWidth(sheetName, colName)
			if width == 0 {
				width = 8.43
			}

			// Determine value
			var cellValue string
			if val, ok := b.engine.Book.RememberedCellValue(sheetName, rowIdx, colIdx); ok {
				cellValue = workbook.DisplayValue(val)
			} else {
				// Check if the cell has a formula in the file, and recalculate its value
				if formula, err := file.GetCellFormula(sheetName, cellName); err == nil && formula != "" {
					if calcVal, err := file.CalcCellValue(sheetName, cellName); err == nil {
						cellValue = calcVal
					}
				}
				if cellValue == "" {
					if rowIdx < len(sheet.Rows) && colIdx < len(sheet.Rows[rowIdx]) {
						cellValue = sheet.Rows[rowIdx][colIdx]
					} else if rowIdx < len(rows) && colIdx < len(rows[rowIdx]) {
						cellValue = rows[rowIdx][colIdx]
					}
				}
			}

			// Merge info
			var rowSpan, colSpan int
			var isMerged, isOverlaid bool
			if info, ok := mergeMap[cellName]; ok {
				rowSpan = info.rowSpan
				colSpan = info.colSpan
				isMerged = info.isMerged
				isOverlaid = info.isOverlaid
			}

			cellView := CellView{
				Coordinate: cellName,
				Value:      cellValue,
				Width:      width,
				Height:     height,
				RowSpan:    rowSpan,
				ColSpan:    colSpan,
				IsMerged:   isMerged,
				IsOverlaid: isOverlaid,
			}

			// Get style information
			styleID, err := file.GetCellStyle(sheetName, cellName)
			if err == nil && styleID > 0 {
				style, err := file.GetStyle(styleID)
				if err == nil && style != nil {
					cellStyle := &CellStyle{}
					if style.Font != nil {
						cellStyle.FontName = style.Font.Family
						cellStyle.FontSize = style.Font.Size
						cellStyle.Bold = style.Font.Bold
						cellStyle.Italic = style.Font.Italic
						cellStyle.FontColor = style.Font.Color
					}
					if len(style.Fill.Color) > 0 {
						cellStyle.FillColor = style.Fill.Color[0]
					}
					if style.Alignment != nil {
						cellStyle.AlignHoriz = style.Alignment.Horizontal
						cellStyle.AlignVert = style.Alignment.Vertical
					}
					cellView.Style = cellStyle
				}
			}

			view[rowIdx][colIdx] = cellView
		}
	}

	return view, nil
}

func (b *Book) getDefaultSheetView(sheet *workbook.Sheet) [][]CellView {
	maxRows := len(sheet.Rows)
	maxCols := 0
	for _, row := range sheet.Rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	view := make([][]CellView, maxRows)
	for rowIdx := 0; rowIdx < maxRows; rowIdx++ {
		view[rowIdx] = make([]CellView, maxCols)
		for colIdx := 0; colIdx < maxCols; colIdx++ {
			cellName, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
			var cellValue string
			if val, ok := b.engine.Book.RememberedCellValue(sheet.Name, rowIdx, colIdx); ok {
				cellValue = workbook.DisplayValue(val)
			} else if colIdx < len(sheet.Rows[rowIdx]) {
				cellValue = sheet.Rows[rowIdx][colIdx]
			}
			view[rowIdx][colIdx] = CellView{
				Coordinate: cellName,
				Value:      cellValue,
				Width:      8.43,
				Height:     15.0,
			}
		}
	}
	return view
}
