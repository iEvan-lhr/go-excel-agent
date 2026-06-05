package engine

import (
	"context"
	"encoding/json"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
	"testing"
)

func TestEngineUpdateCellAndFind(t *testing.T) {
	e := New()
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", "1,200.00"},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}

	found, err := e.Find(context.Background(), FindRequest{
		Sheet:        "Data",
		Type:         "search",
		Query:        "标准键盘",
		SearchColumn: "品名",
	})
	if err != nil {
		t.Fatalf("find failed: %v", err)
	}
	if len(found.([]workbook.FindResult)) == 0 {
		t.Fatal("expected find result")
	}
}

func TestEngineBatchUpdateMultiply(t *testing.T) {
	e := New()
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", "1,200.00"},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}

	diff, err := e.BatchUpdate(context.Background(), BatchUpdateRequest{
		Sheet: "Data",
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
		t.Fatalf("unexpected changed cells: %d", diff.ChangedCells)
	}
	got := e.Book.SheetByName("Data").Rows[1][1]
	if got != "1080" {
		t.Fatalf("unexpected value: %q", got)
	}
}

func TestEngineExecuteUpdateCell(t *testing.T) {
	e := New()
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", "1,200.00"},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}

	_, diff, err := e.Execute(context.Background(), Command{
		Op: "update_cell",
		Target: Target{
			Sheet: "Data",
			Cell:  "B2",
		},
		Args: UpdateCellArgs{Value: 1080},
	})
	if err != nil {
		t.Fatalf("execute failed: %v", err)
	}
	if diff.ChangedCells != 1 {
		t.Fatalf("unexpected changed cells: %d", diff.ChangedCells)
	}
	if got := e.Book.SheetByName("Data").Rows[1][1]; got != "1080" {
		t.Fatalf("unexpected value: %q", got)
	}
}

func TestEngineExecuteJSONCommand(t *testing.T) {
	e := New()
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", "1,200.00"},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}

	var cmd Command
	if err := json.Unmarshal([]byte(`{
		"op": "batch_update",
		"target": {
			"sheet": "Data",
			"column": "单价",
			"scope": {
				"type": "search",
				"query": "标准键盘",
				"search_column": "品名"
			}
		},
		"args": {
			"action": "multiply",
			"value": 0.9
		}
	}`), &cmd); err != nil {
		t.Fatalf("unmarshal failed: %v", err)
	}

	_, diff, err := e.Execute(context.Background(), cmd)
	if err != nil {
		t.Fatalf("execute failed: %v", err)
	}
	if diff.ChangedCells != 1 {
		t.Fatalf("unexpected changed cells: %d", diff.ChangedCells)
	}
	if got := e.Book.SheetByName("Data").Rows[1][1]; got != "1080" {
		t.Fatalf("unexpected value: %q", got)
	}
}
