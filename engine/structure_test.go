package engine

import (
	"context"
	"testing"
)

func TestEngineCreateSheetRecordsStructureDiff(t *testing.T) {
	e := New()
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"Data": {
			{"品名", "单价"},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}

	_, diff, err := e.Execute(context.Background(), Command{
		Op:     "create_sheet",
		Target: Target{Sheet: "销售汇总"},
	})
	if err != nil {
		t.Fatalf("execute failed: %v", err)
	}
	if e.Book.SheetByName("销售汇总") == nil {
		t.Fatal("expected created sheet")
	}
	if len(diff.StructureChanges) != 1 || diff.StructureChanges[0].Type != "sheet_created" {
		t.Fatalf("unexpected diff: %#v", diff)
	}
}

func TestEngineClearCell(t *testing.T) {
	e := New()
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", 1200},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}

	_, diff, err := e.Execute(context.Background(), Command{
		Op:     "clear_cell",
		Target: Target{Sheet: "Data", Cell: "B2"},
	})
	if err != nil {
		t.Fatalf("execute failed: %v", err)
	}
	if got := e.Book.SheetByName("Data").Cell(1, 1); got != "" {
		t.Fatalf("expected empty cell, got %q", got)
	}
	if diff.ChangedCells != 1 {
		t.Fatalf("unexpected diff: %#v", diff)
	}
}

func TestEngineInsertCellsShiftRight(t *testing.T) {
	e := New()
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"Data": {
			{"A", "B", "C"},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}

	_, diff, err := e.Execute(context.Background(), Command{
		Op:     "insert_cells",
		Target: Target{Sheet: "Data", Cell: "B1"},
		Args:   InsertCellsArgs{Shift: "right"},
	})
	if err != nil {
		t.Fatalf("execute failed: %v", err)
	}
	row := e.Book.SheetByName("Data").Rows[0]
	if len(row) != 4 || row[1] != "" || row[2] != "B" || row[3] != "C" {
		t.Fatalf("unexpected shifted row: %#v", row)
	}
	if len(diff.StructureChanges) != 1 || diff.StructureChanges[0].Type != "cells_inserted" {
		t.Fatalf("unexpected diff: %#v", diff)
	}
}
