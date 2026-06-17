package excelizeadapter

import (
	"github.com/iEvan-lhr/go-excel-agent/workbook"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestAdapterSaveAsPreservesStyleAndWritesTypedValue(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.xlsx")
	outputPath := filepath.Join(tempDir, "output.xlsx")

	file := excelize.NewFile()
	if err := file.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatalf("rename sheet failed: %v", err)
	}
	if err := file.SetCellStr("Data", "A1", "品名"); err != nil {
		t.Fatalf("write A1 failed: %v", err)
	}
	if err := file.SetCellStr("Data", "B1", "单价"); err != nil {
		t.Fatalf("write B1 failed: %v", err)
	}
	if err := file.SetCellStr("Data", "A2", "标准键盘"); err != nil {
		t.Fatalf("write A2 failed: %v", err)
	}
	if err := file.SetCellFloat("Data", "B2", 1200, -1, 64); err != nil {
		t.Fatalf("write B2 failed: %v", err)
	}
	styleID, err := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Color: "9C0006"},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"FFF2CC"}, Pattern: 1},
	})
	if err != nil {
		t.Fatalf("create style failed: %v", err)
	}
	if err := file.SetCellStyle("Data", "B2", "B2", styleID); err != nil {
		t.Fatalf("set style failed: %v", err)
	}
	if err := file.SaveAs(inputPath); err != nil {
		t.Fatalf("save input failed: %v", err)
	}
	if err := file.Close(); err != nil {
		t.Fatalf("close input failed: %v", err)
	}

	source, err := excelize.OpenFile(inputPath)
	if err != nil {
		t.Fatalf("open source failed: %v", err)
	}
	sourceStyle, err := source.GetCellStyle("Data", "B2")
	if err != nil {
		t.Fatalf("get source style failed: %v", err)
	}
	_ = source.Close()

	adapter := New()
	book, err := adapter.Open(inputPath)
	if err != nil {
		t.Fatalf("adapter open failed: %v", err)
	}
	sheet := book.SheetByName("Data")
	sheet.SetCell(1, 1, workbook.DisplayValue(1080))
	book.RememberCellValue("Data", 1, 1, 1080)

	result, err := adapter.SaveAs(book, outputPath)
	if err != nil {
		t.Fatalf("save as failed: %v", err)
	}
	if result.ChangedCells != 1 {
		t.Fatalf("unexpected changed cells: %d", result.ChangedCells)
	}

	output, err := excelize.OpenFile(outputPath)
	if err != nil {
		t.Fatalf("open output failed: %v", err)
	}
	defer output.Close()

	value, err := output.GetCellValue("Data", "B2")
	if err != nil {
		t.Fatalf("get value failed: %v", err)
	}
	if value != "1080" {
		t.Fatalf("unexpected value: %q", value)
	}
	cellType, err := output.GetCellType("Data", "B2")
	if err != nil {
		t.Fatalf("get cell type failed: %v", err)
	}
	if cellType != excelize.CellTypeUnset {
		t.Fatalf("expected numeric cell type, got %v", cellType)
	}
	outputStyle, err := output.GetCellStyle("Data", "B2")
	if err != nil {
		t.Fatalf("get output style failed: %v", err)
	}
	if outputStyle != sourceStyle {
		t.Fatalf("style changed, got %d, want %d", outputStyle, sourceStyle)
	}
}

func TestAdapterSaveAsAppliesNewStyleAndFormula(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.xlsx")
	outputPath := filepath.Join(tempDir, "output.xlsx")

	file := excelize.NewFile()
	if err := file.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatalf("rename sheet failed: %v", err)
	}
	if err := file.SetCellStr("Data", "A1", "品名"); err != nil {
		t.Fatalf("write A1 failed: %v", err)
	}
	if err := file.SetCellStr("Data", "B1", "单价"); err != nil {
		t.Fatalf("write B1 failed: %v", err)
	}
	if err := file.SetCellStr("Data", "A2", "标准键盘"); err != nil {
		t.Fatalf("write A2 failed: %v", err)
	}
	if err := file.SetCellFloat("Data", "B2", 1200, -1, 64); err != nil {
		t.Fatalf("write B2 failed: %v", err)
	}
	if err := file.SaveAs(inputPath); err != nil {
		t.Fatalf("save input failed: %v", err)
	}
	_ = file.Close()

	adapter := New()
	book, err := adapter.Open(inputPath)
	if err != nil {
		t.Fatalf("adapter open failed: %v", err)
	}

	book.RememberCellValue("Data", 1, 1, 1000)
	book.RememberCellStyle("Data", "B2", workbook.CellStyle{Bold: true, FontColor: "FF0000"})
	book.RememberFormula("Data", "B3", "=SUM(B2:B2)")

	sheet := book.SheetByName("Data")
	sheet.EnsureSize(2, 1)
	sheet.SetCell(2, 1, "=SUM(B2:B2)")

	result, err := adapter.SaveAs(book, outputPath)
	if err != nil {
		t.Fatalf("save as failed: %v", err)
	}
	if result.ChangedCells < 2 {
		t.Fatalf("unexpected changed cells: %d", result.ChangedCells)
	}

	output, err := excelize.OpenFile(outputPath)
	if err != nil {
		t.Fatalf("open output failed: %v", err)
	}
	defer output.Close()

	formula, err := output.GetCellFormula("Data", "B3")
	if err != nil {
		t.Fatalf("get formula failed: %v", err)
	}
	if formula != "SUM(B2:B2)" {
		t.Fatalf("expected SUM(B2:B2) formula, got: %q", formula)
	}

	styleID, err := output.GetCellStyle("Data", "B2")
	if err != nil {
		t.Fatalf("get style failed: %v", err)
	}
	style, err := output.GetStyle(styleID)
	if err != nil {
		t.Fatalf("get style details failed: %v", err)
	}
	if style.Font == nil || !style.Font.Bold || style.Font.Color != "FF0000" {
		t.Fatalf("style font not applied correctly: %+v", style.Font)
	}
}

func TestAdapterSaveAsAppliesInsertRow(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.xlsx")
	outputPath := filepath.Join(tempDir, "output.xlsx")

	file := excelize.NewFile()
	if err := file.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatalf("rename sheet failed: %v", err)
	}
	if err := file.SetCellStr("Data", "A1", "品名"); err != nil {
		t.Fatalf("write A1 failed: %v", err)
	}
	if err := file.SetCellStr("Data", "B1", "单价"); err != nil {
		t.Fatalf("write B1 failed: %v", err)
	}
	if err := file.SetCellStr("Data", "A2", "标准键盘"); err != nil {
		t.Fatalf("write A2 failed: %v", err)
	}
	if err := file.SetCellFloat("Data", "B2", 1200, -1, 64); err != nil {
		t.Fatalf("write B2 failed: %v", err)
	}
	styleID, err := file.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true}})
	if err != nil {
		t.Fatalf("new style failed: %v", err)
	}
	if err := file.SetCellStyle("Data", "B2", "B2", styleID); err != nil {
		t.Fatalf("set cell style failed: %v", err)
	}
	if err := file.SaveAs(inputPath); err != nil {
		t.Fatalf("save input failed: %v", err)
	}
	_ = file.Close()

	adapter := New()
	book, err := adapter.Open(inputPath)
	if err != nil {
		t.Fatalf("adapter open failed: %v", err)
	}

	rowIdx := 1
	sheet := book.SheetByName("Data")
	sheet.EnsureSize(rowIdx, 0)
	sheet.Rows = append(sheet.Rows, []string{})
	copy(sheet.Rows[rowIdx+1:], sheet.Rows[rowIdx:])
	sheet.Rows[rowIdx] = make([]string, 2)
	sheet.Rows[rowIdx][0] = "新插商品"
	sheet.Rows[rowIdx][1] = "999"

	book.ShiftRowsDown("Data", rowIdx)
	book.RememberCellValue("Data", rowIdx, 0, "新插商品")
	book.RememberCellValue("Data", rowIdx, 1, 999)

	if book.InsertedRows == nil {
		book.InsertedRows = make(map[string][]int)
	}
	book.InsertedRows["Data"] = append(book.InsertedRows["Data"], rowIdx)

	_, err = adapter.SaveAs(book, outputPath)
	if err != nil {
		t.Fatalf("save as failed: %v", err)
	}

	output, err := excelize.OpenFile(outputPath)
	if err != nil {
		t.Fatalf("open output failed: %v", err)
	}
	defer output.Close()

	valB2, _ := output.GetCellValue("Data", "B2")
	if valB2 != "999" {
		t.Fatalf("expected 999 at B2, got %q", valB2)
	}

	valB3, _ := output.GetCellValue("Data", "B3")
	if valB3 != "1200" {
		t.Fatalf("expected 1200 at B3, got %q", valB3)
	}

	styleIDB3, err := output.GetCellStyle("Data", "B3")
	if err != nil {
		t.Fatalf("get style failed: %v", err)
	}
	styleB3, _ := output.GetStyle(styleIDB3)
	if styleB3.Font == nil || !styleB3.Font.Bold {
		t.Fatalf("expected bold font at shifted B3 style, got: %+v", styleB3.Font)
	}
}
