package excelagent

import (
	"context"
	"fmt"
	"path/filepath"
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
