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
