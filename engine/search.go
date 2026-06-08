package engine

import (
	"fmt"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
	"strings"

	"github.com/xuri/excelize/v2"
)

func searchSheet(sheet *workbook.Sheet, query, searchColumn string) ([]workbook.FindResult, error) {
	if sheet == nil {
		return nil, fmt.Errorf("目标 sheet 为空")
	}
	query = strings.TrimSpace(query)

	searchColIdx := -1
	if strings.TrimSpace(searchColumn) != "" {
		if len(sheet.Rows) == 0 {
			return nil, fmt.Errorf("sheet '%s' 没有表头", sheet.Name)
		}
		searchColIdx = findColumnIndexInSheet(sheet, searchColumn)
		if searchColIdx == -1 {
			return nil, fmt.Errorf("找不到搜索列: %s", searchColumn)
		}
	}

	var results []workbook.FindResult
	for rowIdx, row := range sheet.Rows {
		if rowIdx == 0 {
			continue // Skip header row
		}
		if query == "" || query == "*" {
			colIdx := 0
			if searchColIdx != -1 {
				colIdx = searchColIdx
			}
			// If query is "*", it matches any non-empty cell in the search column
			if query == "*" {
				cellVal := ""
				if colIdx < len(row) {
					cellVal = strings.TrimSpace(row[colIdx])
				}
				if cellVal == "" {
					continue
				}
			}
			results = append(results, buildFindResult(sheet, rowIdx, colIdx))
			continue
		}

		if searchColIdx != -1 {
			if searchColIdx < len(row) && strings.Contains(row[searchColIdx], query) {
				results = append(results, buildFindResult(sheet, rowIdx, searchColIdx))
			}
			continue
		}

		for colIdx, cell := range row {
			if strings.Contains(cell, query) {
				results = append(results, buildFindResult(sheet, rowIdx, colIdx))
				break
			}
		}
	}
	return results, nil
}

func buildFindResult(sheet *workbook.Sheet, rowIdx, colIdx int) workbook.FindResult {
	cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
	rowData := []string(nil)
	if rowIdx >= 0 && rowIdx < len(sheet.Rows) {
		rowData = append(rowData, sheet.Rows[rowIdx]...)
	}
	return workbook.FindResult{
		Sheet:    sheet.Name,
		Address:  cell,
		RowIndex: rowIdx + 1,
		ColIndex: colIdx + 1,
		Value:    sheet.Cell(rowIdx, colIdx),
		RowData:  rowData,
	}
}

func findColumnIndex(header []string, name string) int {
	target := strings.TrimSpace(name)
	for i, h := range header {
		if h == name {
			return i
		}
	}
	for i, h := range header {
		if strings.TrimSpace(h) == target {
			return i
		}
	}
	for i, h := range header {
		if strings.EqualFold(strings.TrimSpace(h), target) {
			return i
		}
	}
	return -1
}

func findColumnIndexInSheet(sheet *workbook.Sheet, name string) int {
	if sheet == nil {
		return -1
	}
	target := strings.TrimSpace(name)
	// Scan the first 15 rows to find a matching column header
	for rowIdx := 0; rowIdx < len(sheet.Rows) && rowIdx < 15; rowIdx++ {
		row := sheet.Rows[rowIdx]
		// Try exact match first
		for colIdx, cell := range row {
			if cell == name {
				return colIdx
			}
		}
		// Try trimmed match
		for colIdx, cell := range row {
			if strings.TrimSpace(cell) == target {
				return colIdx
			}
		}
		// Try case-insensitive trimmed match
		for colIdx, cell := range row {
			if strings.EqualFold(strings.TrimSpace(cell), target) {
				return colIdx
			}
		}
	}
	return -1
}
