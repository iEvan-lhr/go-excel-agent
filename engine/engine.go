package engine

import (
	"context"
	"encoding/json"
	"fmt"
	"github.com/iEvan-lhr/go-excel-agent/excelizeadapter"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Engine struct {
	Book    *workbook.Workbook
	Adapter *excelizeadapter.Adapter
}

func New() *Engine {
	return &Engine{
		Book:    workbook.New(),
		Adapter: excelizeadapter.New(),
	}
}

func Open(ctx context.Context, path string) (*Engine, error) {
	e := New()
	if err := e.Open(ctx, path); err != nil {
		return nil, err
	}
	return e, nil
}

func (e *Engine) Open(ctx context.Context, path string) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	book, err := e.Adapter.Open(path)
	if err != nil {
		return err
	}
	e.Book = book
	return nil
}

func (e *Engine) LoadSheets(ctx context.Context, sheets map[string][][]any) error {
	if err := ctx.Err(); err != nil {
		return err
	}

	var loaded []workbook.Sheet
	typedValues := make(map[string]map[string]any)
	for sheetName, rows := range sheets {
		sheet := workbook.Sheet{Name: sheetName}
		for rowIdx, row := range rows {
			stringRow := make([]string, len(row))
			for colIdx, value := range row {
				stringRow[colIdx] = workbook.DisplayValue(value)
				cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
				if err != nil {
					return err
				}
				if typedValues[sheetName] == nil {
					typedValues[sheetName] = make(map[string]any)
				}
				typedValues[sheetName][cell] = value
			}
			sheet.Rows = append(sheet.Rows, stringRow)
		}
		loaded = append(loaded, sheet)
	}

	e.Book = &workbook.Workbook{
		Sheets:      loaded,
		TypedValues: typedValues,
	}
	return nil
}

func (e *Engine) SaveAs(ctx context.Context, path string) (*excelizeadapter.SaveResult, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	return e.Adapter.SaveAs(e.Book, path)
}

func (e *Engine) Execute(ctx context.Context, cmd Command) (any, *workbook.Diff, error) {
	if err := ctx.Err(); err != nil {
		return nil, nil, err
	}

	switch strings.ToLower(strings.TrimSpace(cmd.Op)) {
	case "inspect_workbook":
		return e.Book.Sheets, &workbook.Diff{}, nil
	case "find":
		args, err := decodeCommandArgs[FindArgs](cmd.Args)
		if err != nil {
			return nil, nil, err
		}
		result, err := e.Find(ctx, FindRequest{
			Sheet:        cmd.Target.Sheet,
			Type:         firstNonEmpty(cmd.Target.ScopeType(), args.Type),
			Range:        firstNonEmpty(cmd.Target.Range, args.Range),
			Query:        firstNonEmpty(cmd.Target.SearchQuery, args.Query),
			SearchColumn: firstNonEmpty(cmd.Target.SearchColumn, args.SearchColumn),
		})
		return result, &workbook.Diff{}, err
	case "get_range":
		result, err := e.Find(ctx, FindRequest{Sheet: cmd.Target.Sheet, Type: "range", Range: cmd.Target.Range})
		return result, &workbook.Diff{}, err
	case "update_cell":
		args, err := decodeCommandArgs[UpdateCellArgs](cmd.Args)
		if err != nil {
			return nil, nil, err
		}
		diff, err := e.UpdateCell(ctx, UpdateCellRequest{
			Sheet: cmd.Target.Sheet,
			Cell:  cmd.Target.Cell,
			Value: args.Value,
		})
		return nil, diff, err
	case "clear_cell":
		diff, err := e.ClearCell(ctx, ClearCellRequest{
			Sheet: cmd.Target.Sheet,
			Cell:  cmd.Target.Cell,
		})
		return nil, diff, err
	case "create_sheet":
		args, err := decodeCommandArgs[CreateSheetArgs](cmd.Args)
		if err != nil {
			return nil, nil, err
		}
		diff, err := e.CreateSheet(ctx, CreateSheetRequest{
			Sheet:      cmd.Target.Sheet,
			AfterSheet: args.AfterSheet,
		})
		return nil, diff, err
	case "insert_cells":
		args, err := decodeCommandArgs[InsertCellsArgs](cmd.Args)
		if err != nil {
			return nil, nil, err
		}
		diff, err := e.InsertCells(ctx, InsertCellsRequest{
			Sheet: cmd.Target.Sheet,
			Cell:  cmd.Target.Cell,
			Shift: args.Shift,
		})
		return nil, diff, err
	case "batch_update":
		args, err := decodeCommandArgs[BatchUpdateArgs](cmd.Args)
		if err != nil {
			return nil, nil, err
		}
		scope := Scope{}
		if cmd.Target.Scope != nil {
			scope = *cmd.Target.Scope
		}
		action := UpdateAction{
			Type:    args.Action,
			Value:   args.Value,
			Find:    args.Find,
			Replace: args.Replace,
		}
		diff, err := e.BatchUpdate(ctx, BatchUpdateRequest{
			Sheet:        cmd.Target.Sheet,
			Scope:        scope,
			TargetColumn: cmd.Target.Column,
			Action:       action,
		})
		return nil, diff, err
	case "aggregate":
		args, err := decodeCommandArgs[AggregateArgs](cmd.Args)
		if err != nil {
			return nil, nil, err
		}
		result, err := e.Aggregate(ctx, AggregateRequest{
			Sheet:  cmd.Target.Sheet,
			Column: firstNonEmpty(cmd.Target.Column, args.Column),
			Type:   args.Type,
			Scope:  cmd.Target.Scope,
		})
		return result, &workbook.Diff{}, err
	default:
		return nil, nil, fmt.Errorf("unknown command op: %s", cmd.Op)
	}
}

func decodeCommandArgs[T any](args any) (T, error) {
	var out T
	if args == nil {
		return out, nil
	}
	if typed, ok := args.(T); ok {
		return typed, nil
	}
	if typed, ok := args.(*T); ok {
		return *typed, nil
	}

	raw, err := json.Marshal(args)
	if err != nil {
		return out, fmt.Errorf("无法编码 DSL args: %w", err)
	}
	if err := json.Unmarshal(raw, &out); err != nil {
		return out, fmt.Errorf("无法解析 DSL args: %w", err)
	}
	return out, nil
}

func (t Target) ScopeType() string {
	if t.Scope == nil {
		return ""
	}
	return t.Scope.Type
}

func (e *Engine) Find(ctx context.Context, req FindRequest) (any, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	sheet := e.requireSheet(req.Sheet)
	if sheet == nil {
		return nil, fmt.Errorf("找不到 sheet: %s", req.Sheet)
	}

	switch strings.ToLower(strings.TrimSpace(req.Type)) {
	case "range":
		return getRange(sheet, req.Range)
	case "search":
		return searchSheet(sheet, req.Query, req.SearchColumn)
	default:
		return nil, fmt.Errorf("未知的查找类型: %s", req.Type)
	}
}

func (e *Engine) UpdateCell(ctx context.Context, req UpdateCellRequest) (*workbook.Diff, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	sheet := e.requireSheet(req.Sheet)
	if sheet == nil {
		return nil, fmt.Errorf("找不到 sheet: %s", req.Sheet)
	}

	col, row, err := excelize.CellNameToCoordinates(req.Cell)
	if err != nil {
		return nil, fmt.Errorf("无法识别的单元格地址 '%s': %w", req.Cell, err)
	}
	rowIdx, colIdx := row-1, col-1
	oldValue := sheet.Cell(rowIdx, colIdx)
	newValue := workbook.DisplayValue(req.Value)
	sheet.SetCell(rowIdx, colIdx, newValue)
	e.Book.RememberCellValue(req.Sheet, rowIdx, colIdx, req.Value)

	diff := &workbook.Diff{}
	if oldValue != newValue {
		diff.AddCell(req.Sheet, rowIdx, colIdx, oldValue, newValue)
	}
	return diff, nil
}

func (e *Engine) ClearCell(ctx context.Context, req ClearCellRequest) (*workbook.Diff, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	sheet := e.requireSheet(req.Sheet)
	if sheet == nil {
		return nil, fmt.Errorf("找不到 sheet: %s", req.Sheet)
	}
	col, row, err := excelize.CellNameToCoordinates(req.Cell)
	if err != nil {
		return nil, fmt.Errorf("无法识别的单元格地址 '%s': %w", req.Cell, err)
	}
	rowIdx, colIdx := row-1, col-1
	oldValue := sheet.Cell(rowIdx, colIdx)
	sheet.SetCell(rowIdx, colIdx, "")
	e.Book.ClearCellTypedValue(req.Sheet, rowIdx, colIdx)

	diff := &workbook.Diff{}
	if oldValue != "" {
		diff.AddCell(req.Sheet, rowIdx, colIdx, oldValue, "")
	}
	return diff, nil
}

func (e *Engine) CreateSheet(ctx context.Context, req CreateSheetRequest) (*workbook.Diff, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	if e.Book == nil {
		e.Book = workbook.New()
	}
	if err := e.Book.AddSheet(req.Sheet, req.AfterSheet); err != nil {
		return nil, err
	}
	diff := &workbook.Diff{}
	diff.AddStructure(workbook.StructureChange{
		Type:  "sheet_created",
		Sheet: req.Sheet,
	})
	return diff, nil
}

func (e *Engine) InsertCells(ctx context.Context, req InsertCellsRequest) (*workbook.Diff, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	sheet := e.requireSheet(req.Sheet)
	if sheet == nil {
		return nil, fmt.Errorf("找不到 sheet: %s", req.Sheet)
	}
	col, row, err := excelize.CellNameToCoordinates(req.Cell)
	if err != nil {
		return nil, fmt.Errorf("无法识别的单元格地址 '%s': %w", req.Cell, err)
	}
	rowIdx, colIdx := row-1, col-1
	shift := strings.ToLower(strings.TrimSpace(req.Shift))
	if shift == "" {
		shift = "right"
	}
	switch shift {
	case "right":
		sheet.EnsureSize(rowIdx, colIdx)
		sheet.Rows[rowIdx] = append(sheet.Rows[rowIdx], "")
		copy(sheet.Rows[rowIdx][colIdx+1:], sheet.Rows[rowIdx][colIdx:])
		sheet.Rows[rowIdx][colIdx] = ""
	case "down":
		sheet.EnsureSize(rowIdx, colIdx)
		sheet.Rows = append(sheet.Rows, []string{})
		for r := len(sheet.Rows) - 1; r > rowIdx; r-- {
			sheet.EnsureSize(r, colIdx)
			sheet.EnsureSize(r-1, colIdx)
			sheet.Rows[r][colIdx] = sheet.Rows[r-1][colIdx]
		}
		sheet.Rows[rowIdx][colIdx] = ""
	default:
		return nil, fmt.Errorf("不支持的 insert_cells shift: %s", req.Shift)
	}
	e.Book.ClearSheetTypedValues(req.Sheet)
	diff := &workbook.Diff{}
	diff.AddStructure(workbook.StructureChange{
		Type:  "cells_inserted",
		Sheet: req.Sheet,
		Range: req.Cell,
		Count: 1,
	})
	return diff, nil
}

func (e *Engine) BatchUpdate(ctx context.Context, req BatchUpdateRequest) (*workbook.Diff, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	sheet := e.requireSheet(req.Sheet)
	if sheet == nil {
		return nil, fmt.Errorf("找不到 sheet: %s", req.Sheet)
	}

	rowsToUpdate := make(map[int]bool)
	colsToUpdate := make(map[int]bool)
	if err := e.resolveTargets(sheet, req.Scope, req.TargetColumn, rowsToUpdate, colsToUpdate); err != nil {
		return nil, err
	}
	if len(rowsToUpdate) == 0 || len(colsToUpdate) == 0 {
		return nil, fmt.Errorf("batch_update 未能确定任何要更新的行或列")
	}

	actionType := strings.ToLower(strings.TrimSpace(req.Action.Type))
	switch actionType {
	case "overwrite", "append_suffix", "prepend_prefix", "find_and_replace", "multiply":
	default:
		return nil, fmt.Errorf("不支持的 batch_update action: %s", req.Action.Type)
	}

	diff := &workbook.Diff{}
	for rowIdx := range rowsToUpdate {
		for colIdx := range colsToUpdate {
			sheet.EnsureSize(rowIdx, colIdx)
			oldValue := sheet.Cell(rowIdx, colIdx)
			newValue, typedValue, err := applyAction(oldValue, req.Action)
			if err != nil {
				return nil, err
			}
			sheet.SetCell(rowIdx, colIdx, newValue)
			if newValue != oldValue {
				e.Book.RememberCellValue(req.Sheet, rowIdx, colIdx, typedValue)
				diff.AddCell(req.Sheet, rowIdx, colIdx, oldValue, newValue)
			}
		}
	}
	return diff, nil
}

func (e *Engine) Aggregate(ctx context.Context, req AggregateRequest) (float64, error) {
	if err := ctx.Err(); err != nil {
		return 0, err
	}
	sheet := e.requireSheet(req.Sheet)
	if sheet == nil {
		return 0, fmt.Errorf("找不到 sheet: %s", req.Sheet)
	}
	if len(sheet.Rows) == 0 {
		return 0, nil
	}

	colIdx := findColumnIndexInSheet(sheet, req.Column)
	if colIdx == -1 {
		return 0, fmt.Errorf("找不到列: %s", req.Column)
	}

	rows := make(map[int]bool)
	for rowIdx := 1; rowIdx < len(sheet.Rows); rowIdx++ {
		rows[rowIdx] = true
	}
	if req.Scope != nil {
		rows = make(map[int]bool)
		cols := make(map[int]bool)
		if err := e.resolveTargets(sheet, *req.Scope, req.Column, rows, cols); err != nil {
			return 0, err
		}
	}

	var sum float64
	var count float64
	for rowIdx := range rows {
		value, ok := parseNumber(sheet.Cell(rowIdx, colIdx))
		if ok {
			sum += value
			count++
		}
	}

	switch strings.ToUpper(req.Type) {
	case "SUM":
		return sum, nil
	case "AVERAGE":
		if count == 0 {
			return 0, nil
		}
		return sum / count, nil
	case "COUNT":
		return count, nil
	default:
		return 0, fmt.Errorf("不支持的聚合类型: %s", req.Type)
	}
}

func (e *Engine) requireSheet(name string) *workbook.Sheet {
	if e.Book == nil {
		e.Book = workbook.New()
	}
	return e.Book.SheetByName(name)
}
