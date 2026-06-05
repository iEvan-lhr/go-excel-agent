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
	if query == "" {
		return nil, fmt.Errorf("搜索词不能为空")
	}

	searchColIdx := -1
	if strings.TrimSpace(searchColumn) != "" {
		if len(sheet.Rows) == 0 {
			return nil, fmt.Errorf("sheet '%s' 没有表头", sheet.Name)
		}
		searchColIdx = findColumnIndex(sheet.Rows[0], searchColumn)
		if searchColIdx == -1 {
			return nil, fmt.Errorf("找不到搜索列: %s", searchColumn)
		}
	}

	var results []workbook.FindResult
	for rowIdx, row := range sheet.Rows {
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
