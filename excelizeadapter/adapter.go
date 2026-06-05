package excelizeadapter

import (
	"fmt"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

type Adapter struct {
	CalcMode       string
	FullCalcOnLoad bool
}

type SaveResult struct {
	Mode         string
	ChangedCells int
	Elapsed      time.Duration
}

func New() *Adapter {
	return &Adapter{
		CalcMode:       "auto",
		FullCalcOnLoad: true,
	}
}

func (a *Adapter) Open(path string) (*workbook.Workbook, error) {
	file, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("打开文件 '%s' 失败: %w", path, err)
	}
	defer file.Close()

	var sheets []workbook.Sheet
	for _, sheetName := range file.GetSheetList() {
		rows, err := file.GetRows(sheetName)
		if err != nil {
			return nil, fmt.Errorf("读取 sheet '%s' 失败: %w", sheetName, err)
		}
		sheets = append(sheets, workbook.Sheet{Name: sheetName, Rows: rows})
	}

	book := &workbook.Workbook{Sheets: sheets}
	book.SetSource(path)
	return book, nil
}

func (a *Adapter) SaveAs(book *workbook.Workbook, path string) (*SaveResult, error) {
	if book == nil {
		return nil, fmt.Errorf("workbook is nil")
	}

	start := time.Now()
	if book.SourcePath != "" {
		if _, err := os.Stat(book.SourcePath); err == nil {
			changed, err := a.savePreservingWorkbook(book, path)
			if err != nil {
				return nil, err
			}
			book.MarkSaved(path)
			return &SaveResult{Mode: "preserve", ChangedCells: changed, Elapsed: time.Since(start)}, nil
		}
	}

	if err := a.saveNewWorkbook(book, path); err != nil {
		return nil, err
	}
	book.MarkSaved(path)
	return &SaveResult{Mode: "new_workbook", ChangedCells: workbook.SheetRowsCount(book.Sheets), Elapsed: time.Since(start)}, nil
}

func (a *Adapter) saveNewWorkbook(book *workbook.Workbook, path string) error {
	file := excelize.NewFile()
	defer file.Close()

	for i, sheetData := range book.Sheets {
		if sheetData.Name == "" {
			return fmt.Errorf("工作表名称不能为空")
		}
		if i == 0 {
			if sheetData.Name != "Sheet1" {
				if err := file.SetSheetName("Sheet1", sheetData.Name); err != nil {
					return fmt.Errorf("重命名默认sheet为 '%s' 失败: %w", sheetData.Name, err)
				}
			}
		} else {
			if _, err := file.NewSheet(sheetData.Name); err != nil {
				return fmt.Errorf("创建sheet '%s' 失败: %w", sheetData.Name, err)
			}
		}

		for rowIdx, row := range sheetData.Rows {
			for colIdx, value := range row {
				cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
				if err != nil {
					return err
				}
				valueToWrite := any(value)
				if remembered, ok := book.RememberedCellValue(sheetData.Name, rowIdx, colIdx); ok {
					valueToWrite = remembered
				}
				if err := SetCellAuto(file, sheetData.Name, cell, valueToWrite); err != nil {
					return fmt.Errorf("向sheet '%s' 写入单元格 %s 失败: %w", sheetData.Name, cell, err)
				}
			}
		}
	}

	if err := a.setCalcProps(file); err != nil {
		return err
	}
	if err := file.SaveAs(path); err != nil {
		return fmt.Errorf("保存文件到 '%s' 失败: %w", path, err)
	}
	return nil
}

func (a *Adapter) savePreservingWorkbook(book *workbook.Workbook, path string) (int, error) {
	if len(book.Sheets) == 0 {
		return 0, a.saveNewWorkbook(book, path)
	}

	file, err := excelize.OpenFile(book.SourcePath)
	if err != nil {
		return 0, fmt.Errorf("打开原始文件 '%s' 失败: %w", book.SourcePath, err)
	}
	defer file.Close()

	if err := syncWorkbookSheets(file, book.Sheets); err != nil {
		return 0, err
	}

	originalByName := sheetsByName(book.OriginalSheets)
	changedCells := 0
	for _, sheetData := range book.Sheets {
		written, err := writeSheetValuesPreservingStyles(file, book, sheetData, originalByName[sheetData.Name])
		if err != nil {
			return changedCells, err
		}
		changedCells += written
	}

	if err := a.setCalcProps(file); err != nil {
		return changedCells, err
	}
	if workbook.SamePath(book.SourcePath, path) {
		if err := file.Save(); err != nil {
			return changedCells, fmt.Errorf("保存文件到 '%s' 失败: %w", path, err)
		}
		return changedCells, nil
	}
	if err := file.SaveAs(path); err != nil {
		return changedCells, fmt.Errorf("保存文件到 '%s' 失败: %w", path, err)
	}
	return changedCells, nil
}

func (a *Adapter) setCalcProps(file *excelize.File) error {
	mode := a.CalcMode
	fullCalcOnLoad := a.FullCalcOnLoad
	if err := file.SetCalcProps(&excelize.CalcPropsOptions{
		CalcMode:       &mode,
		FullCalcOnLoad: &fullCalcOnLoad,
	}); err != nil {
		return fmt.Errorf("设置计算属性失败: %w", err)
	}
	return nil
}

func syncWorkbookSheets(file *excelize.File, sheets []workbook.Sheet) error {
	targetSheets := make(map[string]bool, len(sheets))
	existingSheets := make(map[string]bool)
	for _, name := range file.GetSheetList() {
		existingSheets[name] = true
	}

	for _, sheet := range sheets {
		if sheet.Name == "" {
			return fmt.Errorf("工作表名称不能为空")
		}
		targetSheets[sheet.Name] = true
		if !existingSheets[sheet.Name] {
			if _, err := file.NewSheet(sheet.Name); err != nil {
				return fmt.Errorf("创建sheet '%s' 失败: %w", sheet.Name, err)
			}
		}
	}

	for _, sheetName := range file.GetSheetList() {
		if targetSheets[sheetName] {
			continue
		}
		if err := file.DeleteSheet(sheetName); err != nil {
			return fmt.Errorf("删除sheet '%s' 失败: %w", sheetName, err)
		}
	}
	return nil
}

func writeSheetValuesPreservingStyles(file *excelize.File, book *workbook.Workbook, current workbook.Sheet, original *workbook.Sheet) (int, error) {
	maxRows := len(current.Rows)
	if original != nil && len(original.Rows) > maxRows {
		maxRows = len(original.Rows)
	}

	written := 0
	for rowIdx := 0; rowIdx < maxRows; rowIdx++ {
		maxCols := rowLen(current.Rows, rowIdx)
		if original != nil && rowLen(original.Rows, rowIdx) > maxCols {
			maxCols = rowLen(original.Rows, rowIdx)
		}

		for colIdx := 0; colIdx < maxCols; colIdx++ {
			currentValue := current.Cell(rowIdx, colIdx)
			originalValue := ""
			if original != nil {
				originalValue = original.Cell(rowIdx, colIdx)
			}

			valueToWrite := any(currentValue)
			remembered, hasRemembered := book.RememberedCellValue(current.Name, rowIdx, colIdx)
			if hasRemembered {
				valueToWrite = remembered
			}
			if original != nil && currentValue == originalValue && !hasRemembered {
				continue
			}

			cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
			if err != nil {
				return written, err
			}
			if err := SetCellAuto(file, current.Name, cell, valueToWrite); err != nil {
				return written, fmt.Errorf("设置sheet '%s' 单元格 %s 失败: %w", current.Name, cell, err)
			}
			written++
		}
	}
	return written, nil
}

func sheetsByName(sheets []workbook.Sheet) map[string]*workbook.Sheet {
	result := make(map[string]*workbook.Sheet, len(sheets))
	for i := range sheets {
		result[sheets[i].Name] = &sheets[i]
	}
	return result
}

func rowLen(rows [][]string, rowIdx int) int {
	if rowIdx < 0 || rowIdx >= len(rows) {
		return 0
	}
	return len(rows[rowIdx])
}
