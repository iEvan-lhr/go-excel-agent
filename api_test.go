package excelagent

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestPublicAPIWorkbookFlow(t *testing.T) {
	ctx := context.Background()
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "demo.xlsx")
	outputPath := filepath.Join(tempDir, "out.xlsx")

	file := excelize.NewFile()
	if err := file.SetCellStr("Sheet1", "A1", "品名"); err != nil {
		t.Fatalf("write A1 failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "B1", "单价"); err != nil {
		t.Fatalf("write B1 failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "C1", "库存"); err != nil {
		t.Fatalf("write C1 failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "D1", "备注"); err != nil {
		t.Fatalf("write D1 failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "A2", "标准键盘"); err != nil {
		t.Fatalf("write A2 failed: %v", err)
	}
	if err := file.SetCellFloat("Sheet1", "B2", 1200, -1, 64); err != nil {
		t.Fatalf("write B2 failed: %v", err)
	}
	if err := file.SetCellInt("Sheet1", "C2", 8); err != nil {
		t.Fatalf("write C2 failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "D2", "待复核"); err != nil {
		t.Fatalf("write D2 failed: %v", err)
	}
	if err := file.SaveAs(inputPath); err != nil {
		t.Fatalf("save input failed: %v", err)
	}
	if err := file.Close(); err != nil {
		t.Fatalf("close input failed: %v", err)
	}

	book, err := Open(ctx, inputPath)
	if err != nil {
		t.Fatalf("open failed: %v", err)
	}

	summary, err := book.Summary(ctx)
	if err != nil {
		t.Fatalf("summary failed: %v", err)
	}
	if len(summary.Sheets) != 1 || summary.Sheets[0].Name != "Sheet1" {
		t.Fatalf("unexpected summary: %#v", summary)
	}
	if summary.Sheets[0].RowCount != 2 || summary.Sheets[0].ColumnCount != 4 {
		t.Fatalf("unexpected sheet shape: %#v", summary.Sheets[0])
	}

	found, err := book.Find(ctx, FindRequest{
		Sheet:        "Sheet1",
		Type:         "search",
		Query:        "标准键盘",
		SearchColumn: "品名",
	})
	if err != nil {
		t.Fatalf("find failed: %v", err)
	}
	if len(found) != 1 || found[0].Address != "A2" {
		t.Fatalf("unexpected find result: %#v", found)
	}

	rows, err := book.GetRange(ctx, RangeRequest{Sheet: "Sheet1", Range: "A1:B2"})
	if err != nil {
		t.Fatalf("get range failed: %v", err)
	}
	if len(rows) != 2 || len(rows[1]) != 2 || rows[1][0] != "标准键盘" {
		t.Fatalf("unexpected range result: %#v", rows)
	}

	diff, err := book.UpdateCell(ctx, UpdateCellRequest{
		Sheet: "Sheet1",
		Cell:  "D2",
		Value: "已复核",
	})
	if err != nil {
		t.Fatalf("update cell failed: %v", err)
	}
	if diff.ChangedCells != 1 {
		t.Fatalf("unexpected update diff: %#v", diff)
	}

	diff, err = book.BatchUpdate(ctx, BatchUpdateRequest{
		Sheet: "Sheet1",
		Scope: Scope{
			Type:         "search",
			Query:        "标准键盘",
			SearchColumn: "品名",
		},
		TargetColumn: "单价",
		Action: UpdateAction{
			Type:  "multiply",
			Value: 0.9,
		},
	})
	if err != nil {
		t.Fatalf("batch update failed: %v", err)
	}
	if diff.ChangedCells != 1 {
		t.Fatalf("unexpected batch diff: %#v", diff)
	}

	avg, err := book.Aggregate(ctx, AggregateRequest{
		Sheet:  "Sheet1",
		Column: "单价",
		Type:   "AVERAGE",
	})
	if err != nil {
		t.Fatalf("aggregate failed: %v", err)
	}
	if avg != 1080 {
		t.Fatalf("unexpected aggregate: %v", avg)
	}

	if err := book.SaveAs(ctx, outputPath); err != nil {
		t.Fatalf("save as failed: %v", err)
	}

	output, err := excelize.OpenFile(outputPath)
	if err != nil {
		t.Fatalf("open output failed: %v", err)
	}
	defer output.Close()

	value, err := output.GetCellValue("Sheet1", "B2")
	if err != nil {
		t.Fatalf("get B2 failed: %v", err)
	}
	if value != "1080" {
		t.Fatalf("unexpected saved value: %q", value)
	}
}

func TestPublicAPIMemoryFlow(t *testing.T) {
	ctx := context.Background()
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "memory-demo.xlsx")

	file := excelize.NewFile()
	if err := file.SetCellStr("Sheet1", "A1", "品名"); err != nil {
		t.Fatalf("write A1 failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "B1", "单价"); err != nil {
		t.Fatalf("write B1 failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "C1", "库存"); err != nil {
		t.Fatalf("write C1 failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "A2", "标准键盘"); err != nil {
		t.Fatalf("write A2 failed: %v", err)
	}
	if err := file.SetCellFloat("Sheet1", "B2", 1200, -1, 64); err != nil {
		t.Fatalf("write B2 failed: %v", err)
	}
	if err := file.SetCellInt("Sheet1", "C2", 8); err != nil {
		t.Fatalf("write C2 failed: %v", err)
	}
	for row := 3; row <= 200; row++ {
		if err := file.SetCellStr("Sheet1", "A"+cellRow(row), "通用物品"); err != nil {
			t.Fatalf("write generated name failed: %v", err)
		}
	}
	if err := file.SaveAs(inputPath); err != nil {
		t.Fatalf("save input failed: %v", err)
	}
	if err := file.Close(); err != nil {
		t.Fatalf("close input failed: %v", err)
	}

	book, err := Open(ctx, inputPath)
	if err != nil {
		t.Fatalf("open failed: %v", err)
	}

	capsule, err := book.BuildContextCapsule(ctx, ContextRequest{
		Purpose:  PurposePlanUpdate,
		Query:    "标准键盘 单价",
		MaxRows:  1,
		MaxCells: 4,
		MaxNodes: 8,
	})
	if err != nil {
		t.Fatalf("build context capsule failed: %v", err)
	}
	if len(capsule.EvidenceRows) != 1 {
		t.Fatalf("expected bounded evidence rows, got %d", len(capsule.EvidenceRows))
	}

	_, diff, record, err := book.ExecuteAndRemember(ctx, "把标准键盘的单价改为 1080", Command{
		Op: "update_cell",
		Target: Target{
			Sheet: "Sheet1",
			Cell:  "B2",
		},
		Args: UpdateCellArgs{Value: 1080},
	})
	if err != nil {
		t.Fatalf("execute and remember failed: %v", err)
	}
	if diff.ChangedCells != 1 {
		t.Fatalf("unexpected diff: %#v", diff)
	}
	if record.OperationID == "" || len(record.CommandJSON) == 0 {
		t.Fatalf("operation not recorded: %#v", record)
	}
	if book.Memory().Session.LastOperationID != record.OperationID {
		t.Fatalf("session focus not updated: %#v", book.Memory().Session)
	}
}

func TestPublicAPIMemoryOptions(t *testing.T) {
	ctx := context.Background()
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "memory-options.xlsx")

	file := excelize.NewFile()
	if err := file.SetCellStr("Sheet1", "A1", "Header"); err != nil {
		t.Fatalf("write A1 failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "A2", "Value"); err != nil {
		t.Fatalf("write A2 failed: %v", err)
	}
	if err := file.SaveAs(inputPath); err != nil {
		t.Fatalf("save input failed: %v", err)
	}
	if err := file.Close(); err != nil {
		t.Fatalf("close input failed: %v", err)
	}

	book, err := OpenWithMemoryOptions(ctx, inputPath, WithColumnTagger(ColumnTaggerFunc(func(ctx context.Context, req ColumnTagRequest) ([]string, error) {
		return []string{"external_tag"}, nil
	})))
	if err != nil {
		t.Fatalf("open failed: %v", err)
	}
	artifact := book.Memory().Artifacts["current"]
	tags := artifact.Sheets[0].Columns[0].SemanticTags
	if len(tags) != 1 || tags[0] != "external_tag" {
		t.Fatalf("memory option was not applied: %#v", tags)
	}
}

func cellRow(row int) string {
	return fmt.Sprintf("%d", row)
}

func TestGetSheetView(t *testing.T) {
	ctx := context.Background()
	tempDir := t.TempDir()
	filePath := filepath.Join(tempDir, "style-test.xlsx")

	file := excelize.NewFile()
	if err := file.SetSheetName("Sheet1", "Styles"); err != nil {
		t.Fatalf("rename sheet failed: %v", err)
	}
	if err := file.SetCellStr("Styles", "A1", "TestVal"); err != nil {
		t.Fatalf("write A1 failed: %v", err)
	}

	styleID, err := file.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Family: "Arial", Size: 12, Color: "0000FF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFFF00"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	if err != nil {
		t.Fatalf("create style failed: %v", err)
	}
	if err := file.SetCellStyle("Styles", "A1", "A1", styleID); err != nil {
		t.Fatalf("set cell style failed: %v", err)
	}

	if err := file.SetColWidth("Styles", "A", "A", 20); err != nil {
		t.Fatalf("set col width failed: %v", err)
	}
	if err := file.SetRowHeight("Styles", 1, 30); err != nil {
		t.Fatalf("set row height failed: %v", err)
	}

	if err := file.MergeCell("Styles", "A1", "B1"); err != nil {
		t.Fatalf("merge cell failed: %v", err)
	}

	if err := file.SaveAs(filePath); err != nil {
		t.Fatalf("save failed: %v", err)
	}
	_ = file.Close()

	book, err := Open(ctx, filePath)
	if err != nil {
		t.Fatalf("open failed: %v", err)
	}

	view, err := book.GetSheetView(ctx, "Styles")
	if err != nil {
		t.Fatalf("GetSheetView failed: %v", err)
	}

	if len(view) == 0 || len(view[0]) == 0 {
		t.Fatalf("expected non-empty view, got row count %d", len(view))
	}

	cell := view[0][0]
	if cell.Coordinate != "A1" || cell.Value != "TestVal" {
		t.Fatalf("unexpected cell value/coordinate: %+v", cell)
	}

	if cell.ColSpan != 2 || cell.RowSpan != 1 || !cell.IsMerged || cell.IsOverlaid {
		t.Fatalf("unexpected A1 merge properties: colSpan=%d, rowSpan=%d, isMerged=%t, isOverlaid=%t", cell.ColSpan, cell.RowSpan, cell.IsMerged, cell.IsOverlaid)
	}

	cellB1 := view[0][1]
	if !cellB1.IsMerged || !cellB1.IsOverlaid {
		t.Fatalf("expected B1 to be overlaid: isMerged=%t, isOverlaid=%t", cellB1.IsMerged, cellB1.IsOverlaid)
	}

	if cell.Width != 20 || cell.Height != 30 {
		t.Fatalf("unexpected width/height: width=%f, height=%f", cell.Width, cell.Height)
	}

	if cell.Style == nil {
		t.Fatalf("expected style to be set, got nil")
	}

	if cell.Style.FontName != "Arial" || cell.Style.FontSize != 12 || !cell.Style.Bold {
		t.Fatalf("unexpected font details: %+v", cell.Style)
	}

	if cell.Style.FillColor != "FFFF00" {
		t.Fatalf("unexpected fill color: %s", cell.Style.FillColor)
	}

	if cell.Style.AlignHoriz != "center" || cell.Style.AlignVert != "center" {
		t.Fatalf("unexpected alignment: horiz=%s, vert=%s", cell.Style.AlignHoriz, cell.Style.AlignVert)
	}
}

func TestBookExportMarkdownFlow(t *testing.T) {
	ctx := context.Background()
	tempDir := t.TempDir()
	excelPath := filepath.Join(tempDir, "source.xlsx")

	file := excelize.NewFile()
	if err := file.SetCellStr("Sheet1", "A1", "Name"); err != nil {
		t.Fatalf("set cell failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "B1", "Age"); err != nil {
		t.Fatalf("set cell failed: %v", err)
	}
	if err := file.SetCellStr("Sheet1", "A2", "Alice"); err != nil {
		t.Fatalf("set cell failed: %v", err)
	}
	if err := file.SetCellInt("Sheet1", "B2", 30); err != nil {
		t.Fatalf("set cell failed: %v", err)
	}

	// Create another sheet
	_, err := file.NewSheet("Sheet2")
	if err != nil {
		t.Fatalf("new sheet failed: %v", err)
	}
	if err := file.SetCellStr("Sheet2", "A1", "Item"); err != nil {
		t.Fatalf("set cell failed: %v", err)
	}
	if err := file.SetCellStr("Sheet2", "A2", "Book"); err != nil {
		t.Fatalf("set cell failed: %v", err)
	}

	if err := file.SaveAs(excelPath); err != nil {
		t.Fatalf("save excel failed: %v", err)
	}
	_ = file.Close()

	// 1. Test using direct API Book.ExportMarkdown
	book, err := Open(ctx, excelPath)
	if err != nil {
		t.Fatalf("open book failed: %v", err)
	}

	outputDir1 := filepath.Join(tempDir, "md_out_1")
	err = book.ExportMarkdown(ctx, outputDir1)
	if err != nil {
		t.Fatalf("ExportMarkdown failed: %v", err)
	}

	// Check if markdown files are written correctly
	md1Path := filepath.Join(outputDir1, "Sheet1.md")
	content1, err := os.ReadFile(md1Path)
	if err != nil {
		t.Fatalf("read Sheet1.md failed: %v", err)
	}
	if !strings.Contains(string(content1), "| Name | Age |") || !strings.Contains(string(content1), "| Alice | 30 |") {
		t.Errorf("unexpected content in Sheet1.md:\n%s", string(content1))
	}

	md2Path := filepath.Join(outputDir1, "Sheet2.md")
	content2, err := os.ReadFile(md2Path)
	if err != nil {
		t.Fatalf("read Sheet2.md failed: %v", err)
	}
	if !strings.Contains(string(content2), "| Item |") || !strings.Contains(string(content2), "| Book |") {
		t.Errorf("unexpected content in Sheet2.md:\n%s", string(content2))
	}

	// 2. Test using DSL Book.Execute
	outputDir2 := filepath.Join(tempDir, "md_out_2")
	cmd := Command{
		Op: "export_markdown",
		Args: ExportMarkdownArgs{
			OutputDir: outputDir2,
		},
	}
	_, _, err = book.Execute(ctx, cmd)
	if err != nil {
		t.Fatalf("Execute export_markdown failed: %v", err)
	}

	// Verify the files exist and contain content in outputDir2
	if _, err := os.Stat(filepath.Join(outputDir2, "Sheet1.md")); err != nil {
		t.Fatalf("expected Sheet1.md to exist in outputDir2: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outputDir2, "Sheet2.md")); err != nil {
		t.Fatalf("expected Sheet2.md to exist in outputDir2: %v", err)
	}
}
