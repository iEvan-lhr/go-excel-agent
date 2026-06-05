package excelagent

import (
	"context"
	"fmt"

	"github.com/iEvan-lhr/go-excel-agent/engine"
	"github.com/iEvan-lhr/go-excel-agent/excelizeadapter"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
)

type FindRequest = engine.FindRequest
type UpdateCellRequest = engine.UpdateCellRequest
type BatchUpdateRequest = engine.BatchUpdateRequest
type UpdateAction = engine.UpdateAction
type AggregateRequest = engine.AggregateRequest
type Scope = engine.Scope

type Command = engine.Command
type Target = engine.Target
type FindArgs = engine.FindArgs
type UpdateCellArgs = engine.UpdateCellArgs
type BatchUpdateArgs = engine.BatchUpdateArgs
type AggregateArgs = engine.AggregateArgs

type FindResult = workbook.FindResult
type Diff = workbook.Diff
type CellChange = workbook.CellChange
type SaveResult = excelizeadapter.SaveResult

type RangeRequest struct {
	Sheet string `json:"sheet,omitempty"`
	Range string `json:"range,omitempty"`
}

type WorkbookSummary struct {
	Sheets []SheetSummary `json:"sheets"`
}

type SheetSummary struct {
	Name        string   `json:"name"`
	RowCount    int      `json:"row_count"`
	ColumnCount int      `json:"column_count"`
	Headers     []string `json:"headers,omitempty"`
}

type Book struct {
	engine *engine.Engine
}

func New() *Book {
	return &Book{engine: engine.New()}
}

func Open(ctx context.Context, path string) (*Book, error) {
	e, err := engine.Open(ctx, path)
	if err != nil {
		return nil, err
	}
	return &Book{engine: e}, nil
}

func (b *Book) Summary(ctx context.Context) (WorkbookSummary, error) {
	if err := ctx.Err(); err != nil {
		return WorkbookSummary{}, err
	}
	if err := b.ensureEngine(); err != nil {
		return WorkbookSummary{}, err
	}

	summary := WorkbookSummary{Sheets: make([]SheetSummary, 0, len(b.engine.Book.Sheets))}
	for _, sheet := range b.engine.Book.Sheets {
		item := SheetSummary{
			Name:        sheet.Name,
			RowCount:    len(sheet.Rows),
			ColumnCount: workbook.MaxColumnCount(sheet.Rows),
		}
		if len(sheet.Rows) > 0 {
			item.Headers = append([]string(nil), sheet.Rows[0]...)
		}
		summary.Sheets = append(summary.Sheets, item)
	}
	return summary, nil
}

func (b *Book) Find(ctx context.Context, req FindRequest) ([]FindResult, error) {
	if err := b.ensureEngine(); err != nil {
		return nil, err
	}
	if req.Type == "" {
		req.Type = "search"
	}
	result, err := b.engine.Find(ctx, req)
	if err != nil {
		return nil, err
	}
	found, ok := result.([]workbook.FindResult)
	if !ok {
		return nil, fmt.Errorf("find expected search results, got %T", result)
	}
	return found, nil
}

func (b *Book) GetRange(ctx context.Context, req RangeRequest) ([][]string, error) {
	if err := b.ensureEngine(); err != nil {
		return nil, err
	}
	result, err := b.engine.Find(ctx, engine.FindRequest{
		Sheet: req.Sheet,
		Type:  "range",
		Range: req.Range,
	})
	if err != nil {
		return nil, err
	}
	rows, ok := result.([][]string)
	if !ok {
		return nil, fmt.Errorf("get range expected rows, got %T", result)
	}
	return rows, nil
}

func (b *Book) UpdateCell(ctx context.Context, req UpdateCellRequest) (*Diff, error) {
	if err := b.ensureEngine(); err != nil {
		return nil, err
	}
	return b.engine.UpdateCell(ctx, req)
}

func (b *Book) BatchUpdate(ctx context.Context, req BatchUpdateRequest) (*Diff, error) {
	if err := b.ensureEngine(); err != nil {
		return nil, err
	}
	return b.engine.BatchUpdate(ctx, req)
}

func (b *Book) Aggregate(ctx context.Context, req AggregateRequest) (float64, error) {
	if err := b.ensureEngine(); err != nil {
		return 0, err
	}
	return b.engine.Aggregate(ctx, req)
}

func (b *Book) Execute(ctx context.Context, cmd Command) (any, *Diff, error) {
	if err := b.ensureEngine(); err != nil {
		return nil, nil, err
	}
	return b.engine.Execute(ctx, cmd)
}

func (b *Book) SaveAs(ctx context.Context, path string) error {
	_, err := b.SaveAsResult(ctx, path)
	return err
}

func (b *Book) SaveAsResult(ctx context.Context, path string) (*SaveResult, error) {
	if err := b.ensureEngine(); err != nil {
		return nil, err
	}
	return b.engine.SaveAs(ctx, path)
}

func (b *Book) ensureEngine() error {
	if b == nil || b.engine == nil || b.engine.Book == nil {
		return fmt.Errorf("excelagent book is not initialized")
	}
	return nil
}
