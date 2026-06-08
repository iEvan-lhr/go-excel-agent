package memory

import (
	"encoding/json"
	"time"

	"github.com/iEvan-lhr/go-excel-agent/engine"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
)

type ArtifactMemory struct {
	WorkbookID    string         `json:"workbook_id"`
	SourcePath    string         `json:"source_path,omitempty"`
	Fingerprint   string         `json:"fingerprint"`
	SheetCount    int            `json:"sheet_count"`
	Sheets        []SheetProfile `json:"sheets"`
	LastIndexedAt time.Time      `json:"last_indexed_at"`
}

type SheetProfile struct {
	Name         string          `json:"name"`
	RowCount     int             `json:"row_count"`
	ColumnCount  int             `json:"column_count"`
	HeaderRow    int             `json:"header_row"`
	TableRegions []TableRegion   `json:"table_regions,omitempty"`
	Columns      []ColumnProfile `json:"columns"`
	TitleRows    []EvidenceRow   `json:"title_rows,omitempty"`
	SampleRows   []EvidenceRow   `json:"sample_rows,omitempty"`
}

type TableRegion struct {
	Address  string `json:"address"`
	StartRow int    `json:"start_row"`
	EndRow   int    `json:"end_row"`
	StartCol int    `json:"start_col"`
	EndCol   int    `json:"end_col"`
}

type ColumnProfile struct {
	Name                  string         `json:"name"`
	Index                 int            `json:"index"`
	Letter                string         `json:"letter"`
	InferredType          string         `json:"inferred_type"`
	SemanticTags          []string       `json:"semantic_tags,omitempty"`
	SampleValues          []string       `json:"sample_values,omitempty"`
	NonEmptyCount         int            `json:"non_empty_count"`
	EmptyRatio            float64        `json:"empty_ratio"`
	DistinctCountEstimate int            `json:"distinct_count_estimate"`
	NumericStats          *NumericStats  `json:"numeric_stats,omitempty"`
	ValueHints            []ValuePattern `json:"value_hints,omitempty"`
}

type NumericStats struct {
	Count float64 `json:"count"`
	Min   float64 `json:"min"`
	Max   float64 `json:"max"`
	Mean  float64 `json:"mean"`
}

type ValuePattern struct {
	Value string `json:"value"`
	Count int    `json:"count"`
}

type DataGraph struct {
	WorkbookID string      `json:"workbook_id"`
	Nodes      []GraphNode `json:"nodes"`
	Edges      []GraphEdge `json:"edges"`
	BuiltAt    time.Time   `json:"built_at"`
}

type GraphNode struct {
	ID       string         `json:"id"`
	Type     string         `json:"type"`
	Label    string         `json:"label"`
	Ref      EvidenceRef    `json:"ref,omitempty"`
	Summary  string         `json:"summary,omitempty"`
	Tags     []string       `json:"tags,omitempty"`
	Metadata map[string]any `json:"metadata,omitempty"`
}

type GraphEdge struct {
	From string `json:"from"`
	To   string `json:"to"`
	Type string `json:"type"`
}

type EvidenceRef struct {
	WorkbookID string `json:"workbook_id,omitempty"`
	Sheet      string `json:"sheet,omitempty"`
	Address    string `json:"address,omitempty"`
	Row        int    `json:"row,omitempty"`
	Column     int    `json:"column,omitempty"`
}

type EvidenceRow struct {
	Sheet          string            `json:"sheet"`
	RowIndex       int               `json:"row_index"`
	AddressRange   string            `json:"address_range,omitempty"`
	Score          float64           `json:"score,omitempty"`
	MatchedColumns []string          `json:"matched_columns,omitempty"`
	Cells          map[string]string `json:"cells"`
}

type OperationRecord struct {
	OperationID       string            `json:"operation_id"`
	RunID             string            `json:"run_id,omitempty"`
	WorkbookID        string            `json:"workbook_id"`
	Timestamp         time.Time         `json:"timestamp"`
	UserRequest       string            `json:"user_request,omitempty"`
	GeneralizedIntent GeneralizedIntent `json:"generalized_intent"`
	Command           engine.Command    `json:"command"`
	CommandJSON       json.RawMessage   `json:"command_json,omitempty"`
	Locations         []EvidenceRef     `json:"locations,omitempty"`
	Diff              *workbook.Diff    `json:"diff,omitempty"`
	ExecutionResult   any               `json:"execution_result,omitempty"`
	SaveResult        any               `json:"save_result,omitempty"`
	Error             string            `json:"error,omitempty"`
	Metadata          map[string]any    `json:"metadata,omitempty"`
}

type GeneralizedIntent struct {
	IntentType string         `json:"intent_type"`
	Locator    string         `json:"locator,omitempty"`
	Action     string         `json:"action,omitempty"`
	Target     string         `json:"target,omitempty"`
	Properties map[string]any `json:"properties,omitempty"`
}

type ExecutionSummary struct {
	SummaryID   string         `json:"summary_id"`
	OperationID string         `json:"operation_id"`
	WorkbookID  string         `json:"workbook_id"`
	Kind        string         `json:"kind"`
	Text        string         `json:"text"`
	Facts       []MemoryFact   `json:"facts,omitempty"`
	Evidence    []EvidenceRef  `json:"evidence,omitempty"`
	Details     map[string]any `json:"details,omitempty"`
	CreatedAt   time.Time      `json:"created_at"`
}

type MemoryFact struct {
	Key   string `json:"key"`
	Value string `json:"value"`
}

type SessionFocus struct {
	WorkbookID           string        `json:"workbook_id,omitempty"`
	LastSheet            string        `json:"last_sheet,omitempty"`
	LastOperationID      string        `json:"last_operation_id,omitempty"`
	LastContextCapsuleID string        `json:"last_context_capsule_id,omitempty"`
	LastChangedCells     []EvidenceRef `json:"last_changed_cells,omitempty"`
	LastRelevantRows     []EvidenceRef `json:"last_relevant_rows,omitempty"`
	UpdatedAt            time.Time     `json:"updated_at,omitempty"`
}

type ContextPurpose string

const (
	PurposeUnderstandFile ContextPurpose = "understand_file"
	PurposeLocateTarget   ContextPurpose = "locate_target"
	PurposePlanUpdate     ContextPurpose = "plan_update"
	PurposeValidate       ContextPurpose = "validate_command"
	PurposeRepair         ContextPurpose = "repair_error"
	PurposeExplainResult  ContextPurpose = "explain_result"
	PurposeFollowup       ContextPurpose = "answer_followup"
)

type ContextRequest struct {
	WorkbookID        string         `json:"workbook_id,omitempty"`
	Purpose           ContextPurpose `json:"purpose,omitempty"`
	Query             string         `json:"query,omitempty"`
	MaxRows           int            `json:"max_rows,omitempty"`
	MaxCells          int            `json:"max_cells,omitempty"`
	MaxNodes          int            `json:"max_nodes,omitempty"`
	IncludeOperations bool           `json:"include_operations,omitempty"`
}

type ContextCapsule struct {
	CapsuleID          string             `json:"capsule_id"`
	WorkbookID         string             `json:"workbook_id"`
	Purpose            ContextPurpose     `json:"purpose"`
	Query              string             `json:"query,omitempty"`
	IncludedNodes      []GraphNode        `json:"included_nodes,omitempty"`
	EvidenceRows       []EvidenceRow      `json:"evidence_rows,omitempty"`
	RelevantOperations []OperationDigest  `json:"relevant_operations,omitempty"`
	ExecutionSummaries []ExecutionSummary `json:"execution_summaries,omitempty"`
	Budget             ContextBudget      `json:"budget"`
	Excluded           []string           `json:"excluded,omitempty"`
	CreatedAt          time.Time          `json:"created_at"`
}

type OperationDigest struct {
	OperationID       string            `json:"operation_id"`
	Timestamp         time.Time         `json:"timestamp"`
	UserRequest       string            `json:"user_request,omitempty"`
	GeneralizedIntent GeneralizedIntent `json:"generalized_intent"`
	Locations         []EvidenceRef     `json:"locations,omitempty"`
	ChangedCells      int               `json:"changed_cells,omitempty"`
	Error             string            `json:"error,omitempty"`
}

type ContextBudget struct {
	MaxRows       int `json:"max_rows"`
	MaxCells      int `json:"max_cells"`
	MaxNodes      int `json:"max_nodes"`
	IncludedRows  int `json:"included_rows"`
	IncludedCells int `json:"included_cells"`
	IncludedNodes int `json:"included_nodes"`
	SkippedRows   int `json:"skipped_rows,omitempty"`
	SkippedNodes  int `json:"skipped_nodes,omitempty"`
}
