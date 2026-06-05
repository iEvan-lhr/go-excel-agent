package engine

import (
	"fmt"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func getRange(sheet *workbook.Sheet, rangeStr string) ([][]string, error) {
	if sheet == nil {
		return nil, fmt.Errorf("目标 sheet 为空")
	}
	normalized := strings.ToUpper(strings.TrimSpace(rangeStr))
	if normalized == "" {
		return nil, fmt.Errorf("range 不能为空")
	}

	colOnlyRegex := regexp.MustCompile(`^[A-Z]+$`)
	colRangeRegex := regexp.MustCompile(`^([A-Z]+):([A-Z]+)$`)
	rowOnlyRegex := regexp.MustCompile(`^\d+$`)
	rowRangeRegex := regexp.MustCompile(`^(\d+):(\d+)$`)
	cellOnlyRegex := regexp.MustCompile(`^[A-Z]+\d+$`)
	cellRangeRegex := regexp.MustCompile(`^([A-Z]+\d+):([A-Z]+\d+)$`)

	var result [][]string
	switch {
	case cellOnlyRegex.MatchString(normalized):
		col, row, err := excelize.CellNameToCoordinates(normalized)
		if err != nil {
			return nil, err
		}
		return [][]string{{sheet.Cell(row-1, col-1)}}, nil
	case cellRangeRegex.MatchString(normalized):
		parts := cellRangeRegex.FindStringSubmatch(normalized)
		startCol, startRow, err := excelize.CellNameToCoordinates(parts[1])
		if err != nil {
			return nil, err
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(parts[2])
		if err != nil {
			return nil, err
		}
		if startCol > endCol {
			startCol, endCol = endCol, startCol
		}
		if startRow > endRow {
			startRow, endRow = endRow, startRow
		}
		for row := startRow; row <= endRow; row++ {
			out := make([]string, 0, endCol-startCol+1)
			for col := startCol; col <= endCol; col++ {
				out = append(out, sheet.Cell(row-1, col-1))
			}
			result = append(result, out)
		}
	case colRangeRegex.MatchString(normalized):
		parts := colRangeRegex.FindStringSubmatch(normalized)
		startCol, _ := excelize.ColumnNameToNumber(parts[1])
		endCol, _ := excelize.ColumnNameToNumber(parts[2])
		if startCol > endCol {
			startCol, endCol = endCol, startCol
		}
		for rowIdx := range sheet.Rows {
			out := make([]string, 0, endCol-startCol+1)
			for col := startCol; col <= endCol; col++ {
				out = append(out, sheet.Cell(rowIdx, col-1))
			}
			result = append(result, out)
		}
	case colOnlyRegex.MatchString(normalized):
		col, _ := excelize.ColumnNameToNumber(normalized)
		for rowIdx := range sheet.Rows {
			result = append(result, []string{sheet.Cell(rowIdx, col-1)})
		}
	case rowRangeRegex.MatchString(normalized):
		parts := rowRangeRegex.FindStringSubmatch(normalized)
		startRow, _ := strconv.Atoi(parts[1])
		endRow, _ := strconv.Atoi(parts[2])
		if startRow > endRow {
			startRow, endRow = endRow, startRow
		}
		for row := startRow; row <= endRow; row++ {
			if row-1 >= 0 && row-1 < len(sheet.Rows) {
				result = append(result, append([]string(nil), sheet.Rows[row-1]...))
			}
		}
	case rowOnlyRegex.MatchString(normalized):
		row, _ := strconv.Atoi(normalized)
		if row-1 >= 0 && row-1 < len(sheet.Rows) {
			result = append(result, append([]string(nil), sheet.Rows[row-1]...))
		}
	default:
		return nil, fmt.Errorf("无法识别的范围格式: %s", rangeStr)
	}
	return result, nil
}

func (e *Engine) resolveTargets(sheet *workbook.Sheet, scope Scope, targetColumn string, rowsToUpdate, colsToUpdate map[int]bool) error {
	hasTargetColumn := strings.TrimSpace(targetColumn) != ""
	if hasTargetColumn {
		if len(sheet.Rows) == 0 {
			return fmt.Errorf("找不到表头，无法定位目标列 '%s'", targetColumn)
		}
		colIdx := findColumnIndex(sheet.Rows[0], targetColumn)
		if colIdx == -1 {
			return fmt.Errorf("找不到目标列 '%s'", targetColumn)
		}
		colsToUpdate[colIdx] = true
	}

	switch strings.ToLower(strings.TrimSpace(scope.Type)) {
	case "range":
		return addRangeTargets(sheet, scope.Range, hasTargetColumn, rowsToUpdate, colsToUpdate)
	case "search":
		results, err := searchSheet(sheet, scope.Query, scope.SearchColumn)
		if err != nil {
			return err
		}
		for _, result := range results {
			rowsToUpdate[result.RowIndex-1] = true
			if !hasTargetColumn {
				colsToUpdate[result.ColIndex-1] = true
			}
		}
		return nil
	default:
		return fmt.Errorf("不支持的 scope 类型: %s", scope.Type)
	}
}

func addRangeTargets(sheet *workbook.Sheet, rangeStr string, hasTargetColumn bool, rowsToUpdate, colsToUpdate map[int]bool) error {
	normalized := strings.ToUpper(strings.TrimSpace(rangeStr))
	if normalized == "" {
		return fmt.Errorf("range scope 缺少 range 参数")
	}

	colOnlyRegex := regexp.MustCompile(`^[A-Z]+$`)
	colRangeRegex := regexp.MustCompile(`^([A-Z]+):([A-Z]+)$`)
	rowOnlyRegex := regexp.MustCompile(`^\d+$`)
	rowRangeRegex := regexp.MustCompile(`^(\d+):(\d+)$`)
	cellOnlyRegex := regexp.MustCompile(`^[A-Z]+\d+$`)
	cellRangeRegex := regexp.MustCompile(`^([A-Z]+\d+):([A-Z]+\d+)$`)

	addRows := func(start, end int) {
		if start > end {
			start, end = end, start
		}
		for row := start; row <= end; row++ {
			if row > 0 {
				rowsToUpdate[row-1] = true
			}
		}
	}
	addCols := func(start, end int) {
		if start > end {
			start, end = end, start
		}
		for col := start; col <= end; col++ {
			if col > 0 {
				colsToUpdate[col-1] = true
			}
		}
	}
	addAllRows := func() {
		for rowIdx := range sheet.Rows {
			rowsToUpdate[rowIdx] = true
		}
	}
	addAllCols := func() {
		for col := 1; col <= workbook.MaxColumnCount(sheet.Rows); col++ {
			colsToUpdate[col-1] = true
		}
	}

	switch {
	case cellOnlyRegex.MatchString(normalized):
		col, row, err := excelize.CellNameToCoordinates(normalized)
		if err != nil {
			return err
		}
		addRows(row, row)
		if !hasTargetColumn {
			addCols(col, col)
		}
	case cellRangeRegex.MatchString(normalized):
		parts := cellRangeRegex.FindStringSubmatch(normalized)
		startCol, startRow, err := excelize.CellNameToCoordinates(parts[1])
		if err != nil {
			return err
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(parts[2])
		if err != nil {
			return err
		}
		addRows(startRow, endRow)
		if !hasTargetColumn {
			addCols(startCol, endCol)
		}
	case colRangeRegex.MatchString(normalized):
		parts := colRangeRegex.FindStringSubmatch(normalized)
		startCol, _ := excelize.ColumnNameToNumber(parts[1])
		endCol, _ := excelize.ColumnNameToNumber(parts[2])
		addAllRows()
		if !hasTargetColumn {
			addCols(startCol, endCol)
		}
	case colOnlyRegex.MatchString(normalized):
		col, _ := excelize.ColumnNameToNumber(normalized)
		addAllRows()
		if !hasTargetColumn {
			addCols(col, col)
		}
	case rowRangeRegex.MatchString(normalized):
		parts := rowRangeRegex.FindStringSubmatch(normalized)
		startRow, _ := strconv.Atoi(parts[1])
		endRow, _ := strconv.Atoi(parts[2])
		addRows(startRow, endRow)
		if !hasTargetColumn {
			addAllCols()
		}
	case rowOnlyRegex.MatchString(normalized):
		row, _ := strconv.Atoi(normalized)
		addRows(row, row)
		if !hasTargetColumn {
			addAllCols()
		}
	default:
		return fmt.Errorf("无法识别的 range scope: %s", rangeStr)
	}
	return nil
}
