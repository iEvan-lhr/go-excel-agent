package excelagent

import (
	"context"
	"fmt"

	"github.com/iEvan-lhr/go-excel-agent/engine"
	"github.com/iEvan-lhr/go-excel-agent/excelizeadapter"
	"github.com/iEvan-lhr/go-excel-agent/memory"
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
type MemoryStore = memory.Store
type ArtifactMemory = memory.ArtifactMemory
type SheetProfile = memory.SheetProfile
type ColumnProfile = memory.ColumnProfile
type DataGraph = memory.DataGraph
type OperationRecord = memory.OperationRecord
type ExecutionSummary = memory.ExecutionSummary
type GeneralizedIntent = memory.GeneralizedIntent
type SessionFocus = memory.SessionFocus
type ContextPurpose = memory.ContextPurpose
type ContextRequest = memory.ContextRequest
type ContextCapsule = memory.ContextCapsule
type MemoryOption = memory.Option
type TextModel = memory.TextModel
type StreamingTextModel = memory.StreamingTextModel
type TextModelFunc = memory.TextModelFunc
type ModelRequest = memory.ModelRequest
type ModelResponse = memory.ModelResponse
type ModelMessage = memory.ModelMessage
type StreamCallback = memory.StreamCallback
type ColumnTagger = memory.ColumnTagger
type ColumnTaggerFunc = memory.ColumnTaggerFunc
type ColumnTagRequest = memory.ColumnTagRequest
type ModelColumnTagger = memory.ModelColumnTagger
type RuleBasedColumnTagger = memory.RuleBasedColumnTagger
type SemanticTagRule = memory.SemanticTagRule
type IntentGeneralizer = memory.IntentGeneralizer
type IntentGeneralizerFunc = memory.IntentGeneralizerFunc
type ExecutionSummarizer = memory.ExecutionSummarizer
type ExecutionSummarizerFunc = memory.ExecutionSummarizerFunc

var WithColumnTagger = memory.WithColumnTagger
var WithIntentGeneralizer = memory.WithIntentGeneralizer
var WithExecutionSummarizer = memory.WithExecutionSummarizer

const (
	PurposeUnderstandFile ContextPurpose = memory.PurposeUnderstandFile
	PurposeLocateTarget   ContextPurpose = memory.PurposeLocateTarget
	PurposePlanUpdate     ContextPurpose = memory.PurposePlanUpdate
	PurposeValidate       ContextPurpose = memory.PurposeValidate
	PurposeRepair         ContextPurpose = memory.PurposeRepair
	PurposeExplainResult  ContextPurpose = memory.PurposeExplainResult
	PurposeFollowup       ContextPurpose = memory.PurposeFollowup
)

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
	memory *memory.Store
}

func New() *Book {
	return NewWithMemoryOptions()
}

func NewWithMemoryOptions(options ...MemoryOption) *Book {
	return &Book{
		engine: engine.New(),
		memory: memory.NewStore(options...),
	}
}

func Open(ctx context.Context, path string) (*Book, error) {
	return OpenWithMemoryOptions(ctx, path)
}

func OpenWithMemoryOptions(ctx context.Context, path string, options ...MemoryOption) (*Book, error) {
	e, err := engine.Open(ctx, path)
	if err != nil {
		return nil, err
	}
	book := &Book{engine: e, memory: memory.NewStore(options...)}
	if _, err := book.memory.IndexWorkbookContext(ctx, "current", e.Book); err != nil {
		return nil, err
	}
	return book, nil
}

// Reload reloads the workbook from its current source path, keeping the session memory.
func (b *Book) Reload(ctx context.Context) error {
	if err := b.ensureEngine(); err != nil {
		return err
	}
	if b.engine.Book.SourcePath == "" {
		return fmt.Errorf("book has no source path to reload from")
	}

	// Open the workbook file again using the engine
	e, err := engine.Open(ctx, b.engine.Book.SourcePath)
	if err != nil {
		return err
	}
	b.engine = e

	// Re-index the workbook context in the existing memory store
	if _, err := b.memory.IndexWorkbookContext(ctx, b.currentWorkbookID(), e.Book); err != nil {
		return err
	}
	return nil
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

func (b *Book) ExecuteAndRemember(ctx context.Context, userRequest string, cmd Command) (any, *Diff, OperationRecord, error) {
	result, diff, err := b.Execute(ctx, cmd)
	record := OperationRecord{
		WorkbookID:  b.currentWorkbookID(),
		UserRequest: userRequest,
		Command:     cmd,
		Diff:        diff,
	}
	if err != nil {
		record.Error = err.Error()
	}
	remembered, _, recordErr := b.RememberOperation(ctx, record)
	if err != nil {
		return result, diff, remembered, err
	}
	return result, diff, remembered, recordErr
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

func (b *Book) Memory() *MemoryStore {
	if b == nil {
		return nil
	}
	if b.memory == nil {
		b.memory = memory.NewStore()
	}
	return b.memory
}

func (b *Book) IndexMemory(ctx context.Context, workbookID string) (ArtifactMemory, error) {
	if err := ctx.Err(); err != nil {
		return ArtifactMemory{}, err
	}
	if err := b.ensureEngine(); err != nil {
		return ArtifactMemory{}, err
	}
	return b.Memory().IndexWorkbookContext(ctx, workbookID, b.engine.Book)
}

func (b *Book) BuildContextCapsule(ctx context.Context, req ContextRequest) (ContextCapsule, error) {
	if err := ctx.Err(); err != nil {
		return ContextCapsule{}, err
	}
	if err := b.ensureEngine(); err != nil {
		return ContextCapsule{}, err
	}
	if req.WorkbookID == "" {
		req.WorkbookID = b.currentWorkbookID()
	}
	return b.Memory().BuildContextCapsuleContext(ctx, b.engine.Book, req)
}

func (b *Book) RememberOperation(ctx context.Context, record OperationRecord) (OperationRecord, ExecutionSummary, error) {
	if err := ctx.Err(); err != nil {
		return record, ExecutionSummary{}, err
	}
	if err := b.ensureEngine(); err != nil {
		return record, ExecutionSummary{}, err
	}
	if record.WorkbookID == "" {
		record.WorkbookID = b.currentWorkbookID()
	}
	return b.Memory().RecordOperationContext(ctx, record)
}

func (b *Book) ensureEngine() error {
	if b == nil || b.engine == nil || b.engine.Book == nil {
		return fmt.Errorf("excelagent book is not initialized")
	}
	return nil
}

func (b *Book) currentWorkbookID() string {
	if b == nil || b.memory == nil || b.memory.Session.WorkbookID == "" {
		return "current"
	}
	return b.memory.Session.WorkbookID
}
