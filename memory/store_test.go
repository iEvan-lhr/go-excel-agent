package memory

import (
	"context"
	"testing"

	"github.com/iEvan-lhr/go-excel-agent/engine"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
)

func TestStoreBuildsContextCapsuleWithoutCarryingAllRows(t *testing.T) {
	book := workbook.FromSheets([]workbook.Sheet{
		{
			Name: "库存台账",
			Rows: append([][]string{
				{"品名", "单价", "库存", "状态"},
				{"标准键盘", "1200", "8", "正常"},
			}, generatedRows(300)...),
		},
	})

	store := NewStore()
	artifact := store.IndexWorkbook("wb1", book)
	if artifact.SheetCount != 1 {
		t.Fatalf("unexpected sheet count: %d", artifact.SheetCount)
	}
	if len(store.Graphs["wb1"].Nodes) == 0 {
		t.Fatal("expected graph nodes")
	}
	if len(artifact.Sheets[0].Columns[0].SemanticTags) != 0 {
		t.Fatalf("default indexing should not hard-code semantic tags: %#v", artifact.Sheets[0].Columns[0].SemanticTags)
	}

	capsule, err := store.BuildContextCapsule(book, ContextRequest{
		WorkbookID: "wb1",
		Purpose:    PurposePlanUpdate,
		Query:      "标准键盘 单价",
		MaxRows:    2,
		MaxCells:   8,
		MaxNodes:   6,
	})
	if err != nil {
		t.Fatalf("build capsule failed: %v", err)
	}
	if len(capsule.EvidenceRows) != 1 {
		t.Fatalf("expected one relevant row, got %d", len(capsule.EvidenceRows))
	}
	if capsule.EvidenceRows[0].Cells["品名"] != "标准键盘" {
		t.Fatalf("unexpected evidence row: %#v", capsule.EvidenceRows[0])
	}
	if capsule.Budget.IncludedRows > capsule.Budget.MaxRows {
		t.Fatalf("row budget exceeded: %#v", capsule.Budget)
	}
	if capsule.Budget.IncludedCells > capsule.Budget.MaxCells {
		t.Fatalf("cell budget exceeded: %#v", capsule.Budget)
	}
}

func TestStoreUsesInjectedColumnTagger(t *testing.T) {
	book := workbook.FromSheets([]workbook.Sheet{
		{
			Name: "Data",
			Rows: [][]string{
				{"Any Header"},
				{"sample"},
			},
		},
	})
	store := NewStore(WithColumnTagger(ColumnTaggerFunc(func(ctx context.Context, req ColumnTagRequest) ([]string, error) {
		if req.ColumnName != "Any Header" {
			t.Fatalf("unexpected tag request: %#v", req)
		}
		return []string{"custom_tag"}, nil
	})))

	artifact, err := store.IndexWorkbookContext(context.Background(), "wb1", book)
	if err != nil {
		t.Fatalf("index failed: %v", err)
	}
	tags := artifact.Sheets[0].Columns[0].SemanticTags
	if len(tags) != 1 || tags[0] != "custom_tag" {
		t.Fatalf("unexpected tags: %#v", tags)
	}
}

func TestStoreDetectsHeaderRowBelowTitleRows(t *testing.T) {
	book := workbook.FromSheets([]workbook.Sheet{
		{
			Name: "报价总览表",
			Rows: [][]string{
				{"项目报价总览"},
				{""},
				{"资源名称", "规格", "数量", "单价"},
				{"设备接入服务", "标准版", "2", "100"},
			},
		},
	})

	store := NewStore()
	artifact := store.IndexWorkbook("wb1", book)
	sheet := artifact.Sheets[0]
	if sheet.HeaderRow != 3 {
		t.Fatalf("unexpected header row: %d", sheet.HeaderRow)
	}
	if sheet.Columns[1].Name != "规格" || sheet.Columns[3].Name != "单价" {
		t.Fatalf("headers were not detected: %#v", sheet.Columns)
	}
	if len(sheet.TitleRows) != 1 || sheet.TitleRows[0].Cells["A"] != "项目报价总览" {
		t.Fatalf("title rows not captured: %#v", sheet.TitleRows)
	}

	capsule, err := store.BuildContextCapsule(book, ContextRequest{
		WorkbookID: "wb1",
		Purpose:    PurposeUnderstandFile,
		MaxRows:    4,
		MaxCells:   16,
	})
	if err != nil {
		t.Fatalf("build capsule failed: %v", err)
	}
	foundHeaderMappedSample := false
	for _, row := range capsule.EvidenceRows {
		if row.Cells["资源名称"] == "设备接入服务" && row.Cells["单价"] == "100" {
			foundHeaderMappedSample = true
			break
		}
	}
	if !foundHeaderMappedSample {
		t.Fatalf("sample row was not mapped with detected headers: %#v", capsule.EvidenceRows)
	}
}

func TestStoreRecordsOperationLedgerAndAdaptiveSummary(t *testing.T) {
	store := NewStore()
	diff := &workbook.Diff{}
	diff.AddCell("库存台账", 1, 3, "待复核", "已复核")

	record, summary, err := store.RecordOperation(OperationRecord{
		WorkbookID:  "wb1",
		UserRequest: "把标准键盘备注改成已复核",
		Command: engine.Command{
			Op: "update_cell",
			Target: engine.Target{
				Sheet: "库存台账",
				Cell:  "D2",
			},
			Args: engine.UpdateCellArgs{Value: "已复核"},
		},
		Diff: diff,
	})
	if err != nil {
		t.Fatalf("record operation failed: %v", err)
	}
	if record.OperationID == "" || len(record.CommandJSON) == 0 {
		t.Fatalf("operation was not normalized: %#v", record)
	}
	if len(record.Locations) != 1 || record.Locations[0].Address != "D2" {
		t.Fatalf("unexpected locations: %#v", record.Locations)
	}
	if summary.Kind != "single_cell_change" {
		t.Fatalf("unexpected summary kind: %s", summary.Kind)
	}
	if store.Session.LastOperationID != record.OperationID {
		t.Fatalf("session focus not updated: %#v", store.Session)
	}
}

func TestStoreUsesInjectedGeneralizerAndSummarizer(t *testing.T) {
	store := NewStore(
		WithIntentGeneralizer(IntentGeneralizerFunc(func(ctx context.Context, record OperationRecord) (GeneralizedIntent, error) {
			return GeneralizedIntent{IntentType: "custom_intent", Action: "custom_action"}, nil
		})),
		WithExecutionSummarizer(ExecutionSummarizerFunc(func(ctx context.Context, record OperationRecord) (ExecutionSummary, error) {
			return ExecutionSummary{
				Kind: "custom_summary",
				Text: "custom model summary",
			}, nil
		})),
	)

	record, summary, err := store.RecordOperationContext(context.Background(), OperationRecord{
		WorkbookID: "wb1",
		Command: engine.Command{
			Op: "update_cell",
			Target: engine.Target{
				Sheet: "Data",
				Cell:  "A1",
			},
			Args: engine.UpdateCellArgs{Value: "new"},
		},
	})
	if err != nil {
		t.Fatalf("record operation failed: %v", err)
	}
	if record.GeneralizedIntent.IntentType != "custom_intent" {
		t.Fatalf("custom generalizer was not used: %#v", record.GeneralizedIntent)
	}
	if summary.Kind != "custom_summary" || summary.Text != "custom model summary" {
		t.Fatalf("custom summarizer was not used: %#v", summary)
	}
	if summary.SummaryID == "" || summary.OperationID != record.OperationID {
		t.Fatalf("summary was not normalized: %#v", summary)
	}
}

func generatedRows(count int) [][]string {
	rows := make([][]string, 0, count)
	for i := 0; i < count; i++ {
		rows = append(rows, []string{
			"通用物品",
			"10",
			"100",
			"正常",
		})
	}
	return rows
}
