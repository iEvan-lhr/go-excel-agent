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

func TestEngineExecuteUpdateStyle(t *testing.T) {
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
		Op: "update_style",
		Target: Target{
			Sheet: "Data",
			Cell:  "A2",
		},
		Args: UpdateStyleArgs{Style: workbook.CellStyle{
			Bold:      true,
			FontColor: "FFFF0000",
		}},
	})
	if err != nil {
		t.Fatalf("execute failed: %v", err)
	}
	if len(diff.StructureChanges) != 1 || diff.StructureChanges[0].Type != "style_updated" {
		t.Fatalf("expected style_updated structure change, got: %+v", diff.StructureChanges)
	}
	style, ok := e.Book.RememberedCellStyle("Data", 1, 0)
	if !ok || !style.Bold || style.FontColor != "FFFF0000" {
		t.Fatalf("style not remembered correctly: %+v", style)
	}
}

func TestEngineExecuteWriteFormula(t *testing.T) {
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
		Op: "write_formula",
		Target: Target{
			Sheet: "Data",
			Cell:  "B3",
		},
		Args: WriteFormulaArgs{Formula: "SUM(B2:B2)"},
	})
	if err != nil {
		t.Fatalf("execute failed: %v", err)
	}
	if len(diff.StructureChanges) != 1 || diff.StructureChanges[0].Type != "formula_updated" {
		t.Fatalf("expected formula_updated structure change, got: %+v", diff.StructureChanges)
	}
	formula, ok := e.Book.RememberedFormula("Data", 2, 1)
	t.Logf("Formulas in Book: %+v, ok: %t, formula: %q", e.Book.Formulas, ok, formula)
	if !ok || formula != "=SUM(B2:B2)" {
		t.Fatalf("formula not remembered correctly: %s", formula)
	}
	got := e.Book.SheetByName("Data").Rows[2][1]
	if got != "=SUM(B2:B2)" {
		t.Fatalf("unexpected row cell value for formula: %q", got)
	}
}

func TestEngineExecuteSequence(t *testing.T) {
	e := New()
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", "1,200.00"},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}

	_, diff, err := e.ExecuteSequence(context.Background(), []Command{
		{
			Op: "update_cell",
			Target: Target{
				Sheet: "Data",
				Cell:  "B2",
			},
			Args: UpdateCellArgs{Value: 1000},
		},
		{
			Op: "update_style",
			Target: Target{
				Sheet: "Data",
				Cell:  "B2",
			},
			Args: UpdateStyleArgs{Style: workbook.CellStyle{Bold: true}},
		},
	})
	if err != nil {
		t.Fatalf("sequence execution failed: %v", err)
	}
	if diff.ChangedCells != 1 {
		t.Fatalf("expected 1 cell change, got: %d", diff.ChangedCells)
	}
	if len(diff.StructureChanges) != 1 || diff.StructureChanges[0].Type != "style_updated" {
		t.Fatalf("expected 1 style_updated structure change, got: %+v", diff.StructureChanges)
	}
}

func TestEngineExecuteInsertRow(t *testing.T) {
	e := New()
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", 1200},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}

	e.Book.RememberCellStyle("Data", "B2", workbook.CellStyle{Bold: true})
	e.Book.RememberCellValue("Data", 1, 1, 1200)

	_, diff, err := e.Execute(context.Background(), Command{
		Op: "insert_row",
		Target: Target{
			Sheet: "Data",
			Cell:  "A2",
		},
		Args: InsertRowArgs{Values: []any{"新插商品", 999}},
	})
	if err != nil {
		t.Fatalf("execute failed: %v", err)
	}
	if len(diff.StructureChanges) != 1 || diff.StructureChanges[0].Type != "row_inserted" {
		t.Fatalf("expected row_inserted structure change, got: %+v", diff.StructureChanges)
	}

	newRow := e.Book.SheetByName("Data").Rows[1]
	if newRow[0] != "新插商品" || newRow[1] != "999" {
		t.Fatalf("unexpected inserted row values: %+v", newRow)
	}

	shiftedRow := e.Book.SheetByName("Data").Rows[2]
	if shiftedRow[0] != "标准键盘" || shiftedRow[1] != "1200" {
		t.Fatalf("unexpected shifted row values: %+v", shiftedRow)
	}

	style, ok := e.Book.RememberedCellStyle("Data", 2, 1)
	if !ok || !style.Bold {
		t.Fatalf("style not shifted correctly: %+v", style)
	}

	val, ok := e.Book.RememberedCellValue("Data", 2, 1)
	if !ok || workbook.DisplayValue(val) != "1200" {
		t.Fatalf("value not shifted correctly: %v", val)
	}
}
