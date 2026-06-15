package excelagent

import (
	"context"
	"fmt"
	"regexp"
	"sort"
	"strings"
	"time"

	"github.com/iEvan-lhr/go-excel-agent/memory"
	"github.com/iEvan-lhr/go-excel-agent/ops"
)

type InteractionState string

const (
	StateReceived          InteractionState = "received"
	StateContextBuilt      InteractionState = "context_built"
	StateVectorRetrieved   InteractionState = "vector_retrieved"
	StateOperationMatched  InteractionState = "operation_matched"
	StateArgsExtracted     InteractionState = "args_extracted"
	StateValidated         InteractionState = "validated"
	StateNeedClarification InteractionState = "need_clarification"
	StateNeedConfirmation  InteractionState = "need_confirmation"
	StateExecutable        InteractionState = "executable"
	StateExecuting         InteractionState = "executing"
	StateExecuted          InteractionState = "executed"
	StateRemembered        InteractionState = "remembered"
	StateFailed            InteractionState = "failed"
)

type OperationCandidate struct {
	Op     string  `json:"op"`
	Score  float64 `json:"score"`
	Source string  `json:"source,omitempty"`
	Kind   string  `json:"kind,omitempty"`
	Level  string  `json:"level,omitempty"`
	Risk   string  `json:"risk,omitempty"`
	Reason string  `json:"reason,omitempty"`
}

type StateEvent struct {
	RunID       string               `json:"run_id"`
	Step        int                  `json:"step"`
	State       InteractionState     `json:"state"`
	Message     string               `json:"message,omitempty"`
	UserRequest string               `json:"user_request,omitempty"`
	Op          string               `json:"op,omitempty"`
	Level       string               `json:"level,omitempty"`
	Risk        string               `json:"risk,omitempty"`
	Candidates  []OperationCandidate `json:"candidates,omitempty"`
	Command     *Command             `json:"command,omitempty"`
	Diff        *Diff                `json:"diff,omitempty"`
	Error       string               `json:"error,omitempty"`
	Metadata    map[string]any       `json:"metadata,omitempty"`
	CreatedAt   time.Time            `json:"created_at"`
}

type StateEventSink interface {
	OnStateEvent(ctx context.Context, event StateEvent) error
}

type StateEventSinkFunc func(ctx context.Context, event StateEvent) error

func (f StateEventSinkFunc) OnStateEvent(ctx context.Context, event StateEvent) error {
	return f(ctx, event)
}

type InteractionRequest struct {
	UserRequest string         `json:"user_request"`
	Command     *Command       `json:"command,omitempty"`
	WorkbookID  string         `json:"workbook_id,omitempty"`
	Purpose     ContextPurpose `json:"purpose,omitempty"`
	Confirmed   bool           `json:"confirmed,omitempty"`
	MaxRows     int            `json:"max_rows,omitempty"`
	MaxCells    int            `json:"max_cells,omitempty"`
	MaxNodes    int            `json:"max_nodes,omitempty"`
}

type InteractionOptions struct {
	Registry                     *ops.Registry
	VectorStore                  memory.VectorStore
	EventSink                    StateEventSink
	RequireConfirmationForLevels []ops.OperationLevel
	VectorLimit                  int
}

type InteractionResult struct {
	RunID      string               `json:"run_id"`
	Status     InteractionState     `json:"status"`
	Message    string               `json:"message,omitempty"`
	Command    *Command             `json:"command,omitempty"`
	Diff       *Diff                `json:"diff,omitempty"`
	Record     OperationRecord      `json:"record,omitempty"`
	Candidates []OperationCandidate `json:"candidates,omitempty"`
	Events     []StateEvent         `json:"events"`
}

func (b *Book) HandleInteraction(ctx context.Context, req InteractionRequest, options InteractionOptions) (InteractionResult, error) {
	if err := b.ensureEngine(); err != nil {
		return InteractionResult{}, err
	}
	registry := options.Registry
	if registry == nil {
		registry = ops.BuiltinRegistry()
	}
	vectorStore := options.VectorStore
	if vectorStore != nil {
		if err := memory.SeedOperationSpecs(ctx, vectorStore, registry.List()); err != nil {
			return InteractionResult{}, err
		}
	}
	run := &interactionRun{
		book:      b,
		req:       req,
		options:   options,
		registry:  registry,
		vector:    vectorStore,
		runID:     fmt.Sprintf("run_%d", time.Now().UnixNano()),
		startTime: time.Now(),
	}
	return run.run(ctx)
}

type interactionRun struct {
	book      *Book
	req       InteractionRequest
	options   InteractionOptions
	registry  *ops.Registry
	vector    memory.VectorStore
	runID     string
	step      int
	events    []StateEvent
	startTime time.Time
}

func (r *interactionRun) run(ctx context.Context) (InteractionResult, error) {
	r.emit(ctx, StateReceived, "received user request", nil)

	purpose := r.req.Purpose
	if purpose == "" {
		purpose = PurposePlanUpdate
	}
	capsule, err := r.book.BuildContextCapsule(ctx, ContextRequest{
		WorkbookID:        r.req.WorkbookID,
		Purpose:           purpose,
		Query:             r.req.UserRequest,
		MaxRows:           r.req.MaxRows,
		MaxCells:          r.req.MaxCells,
		MaxNodes:          r.req.MaxNodes,
		IncludeOperations: true,
	})
	if err != nil {
		return r.fail(ctx, nil, err)
	}
	r.emit(ctx, StateContextBuilt, "context capsule built", map[string]any{
		"capsule_id": capsule.CapsuleID,
		"rows":       capsule.Budget.IncludedRows,
		"nodes":      capsule.Budget.IncludedNodes,
	})

	candidates, err := r.retrieveCandidates(ctx)
	if err != nil {
		return r.fail(ctx, nil, err)
	}
	r.emitWithCandidates(ctx, StateVectorRetrieved, "retrieved operation candidates", candidates, nil, nil)

	cmd := r.req.Command
	if cmd == nil {
		planned, decision := r.planFromCandidates(candidates)
		if decision == "clarify" {
			r.emitWithCandidates(ctx, StateNeedClarification, "operation candidates are ambiguous or parameters are incomplete", candidates, nil, map[string]any{"decision": decision})
			return r.result(StateNeedClarification, "需要更多信息才能确定操作。", nil, nil, OperationRecord{}, candidates), nil
		}
		cmd = planned
	}
	if cmd == nil {
		r.emitWithCandidates(ctx, StateNeedClarification, "no executable command was produced", candidates, nil, nil)
		return r.result(StateNeedClarification, "没有生成可执行指令。", nil, nil, OperationRecord{}, candidates), nil
	}

	spec, ok := r.registry.Get(cmd.Op)
	if !ok {
		return r.fail(ctx, cmd, fmt.Errorf("operation spec not found: %s", cmd.Op))
	}
	r.emitWithCommand(ctx, StateOperationMatched, "matched operation", cmd, map[string]any{
		"level": spec.Level,
		"risk":  spec.Risk,
	})
	r.emitWithCommand(ctx, StateArgsExtracted, "operation arguments extracted", cmd, nil)

	if err := r.validateCommand(*cmd, spec); err != nil {
		r.emitWithCommand(ctx, StateNeedClarification, err.Error(), cmd, nil)
		return r.result(StateNeedClarification, err.Error(), cmd, nil, OperationRecord{}, candidates), nil
	}
	r.emitWithCommand(ctx, StateValidated, "command validated against workbook state", cmd, nil)

	if r.requiresConfirmation(spec) && !r.req.Confirmed {
		message := confirmationMessage(*cmd, spec)
		r.emitWithCommand(ctx, StateNeedConfirmation, message, cmd, nil)
		return r.result(StateNeedConfirmation, message, cmd, nil, OperationRecord{}, candidates), nil
	}

	r.emitWithCommand(ctx, StateExecutable, "command is executable", cmd, nil)
	r.emitWithCommand(ctx, StateExecuting, "executing command", cmd, nil)
	_, diff, record, err := r.book.ExecuteAndRemember(ctx, r.req.UserRequest, *cmd)
	if err != nil {
		if r.vector != nil {
			_ = r.vector.Upsert(ctx, []memory.VectorRecord{memory.VectorRecordFromOperation(record, spec)})
		}
		return r.fail(ctx, cmd, err)
	}
	r.emitWithDiff(ctx, StateExecuted, "command executed", cmd, diff)

	if r.vector != nil {
		experience := memory.VectorRecordFromOperation(record, spec)
		if err := r.vector.Upsert(ctx, []memory.VectorRecord{experience}); err != nil {
			return r.fail(ctx, cmd, err)
		}
		r.emitWithCommand(ctx, StateRemembered, "operation memory and vector experience recorded", cmd, map[string]any{
			"experience_id": experience.ID,
		})
	}
	return r.result(StateRemembered, "执行完成。", cmd, diff, record, candidates), nil
}

func (r *interactionRun) retrieveCandidates(ctx context.Context) ([]OperationCandidate, error) {
	var candidates []OperationCandidate
	if r.vector != nil {
		limit := r.options.VectorLimit
		if limit <= 0 {
			limit = 8
		}
		results, err := r.vector.Search(ctx, memory.VectorSearchRequest{
			Query:      r.req.UserRequest,
			WorkbookID: r.req.WorkbookID,
			Limit:      limit,
		})
		if err != nil {
			return nil, err
		}
		for _, result := range results {
			candidates = append(candidates, OperationCandidate{
				Op:     result.Record.Op,
				Score:  result.Score,
				Source: result.Record.ID,
				Kind:   string(result.Record.Kind),
				Level:  string(result.Record.Level),
				Risk:   string(result.Record.Risk),
				Reason: result.Reason,
			})
		}
	}
	if len(candidates) == 0 {
		candidates = fallbackCandidates(r.req.UserRequest, r.registry)
	}
	sort.Slice(candidates, func(i, j int) bool {
		if candidates[i].Score == candidates[j].Score {
			return candidates[i].Op < candidates[j].Op
		}
		return candidates[i].Score > candidates[j].Score
	})
	return candidates, nil
}

func (r *interactionRun) planFromCandidates(candidates []OperationCandidate) (*Command, string) {
	if len(candidates) == 0 {
		return nil, "clarify"
	}
	top := candidates[0]
	if top.Op == "" {
		return nil, "clarify"
	}
	conflicts := 0
	for _, candidate := range candidates {
		if candidate.Op != "" && candidate.Op != top.Op && candidate.Score >= top.Score*0.8 {
			conflicts++
		}
	}
	if conflicts > 0 {
		return nil, "clarify"
	}
	switch top.Op {
	case "create_sheet":
		sheet := extractSheetName(r.req.UserRequest)
		if sheet == "" {
			return nil, "clarify"
		}
		return &Command{Op: "create_sheet", Target: Target{Sheet: sheet}}, "reuse"
	case "clear_cell":
		cell := extractCellAddress(r.req.UserRequest)
		sheet := r.defaultSheet()
		if sheet == "" || cell == "" {
			return nil, "clarify"
		}
		return &Command{Op: "clear_cell", Target: Target{Sheet: sheet, Cell: cell}}, "reuse"
	case "insert_cells":
		cell := extractCellAddress(r.req.UserRequest)
		sheet := r.defaultSheet()
		if sheet == "" || cell == "" {
			return nil, "clarify"
		}
		return &Command{
			Op:     "insert_cells",
			Target: Target{Sheet: sheet, Cell: cell},
			Args:   InsertCellsArgs{Shift: extractInsertShift(r.req.UserRequest)},
		}, "reuse"
	default:
		return nil, "clarify"
	}
}

func (r *interactionRun) validateCommand(cmd Command, spec ops.OperationSpec) error {
	if strings.TrimSpace(cmd.Op) == "" {
		return fmt.Errorf("command op is required")
	}
	switch strings.ToLower(strings.TrimSpace(cmd.Op)) {
	case "create_sheet":
		if strings.TrimSpace(cmd.Target.Sheet) == "" {
			return fmt.Errorf("create_sheet requires target.sheet")
		}
		if r.book.engine.Book.SheetByName(cmd.Target.Sheet) != nil {
			return fmt.Errorf("sheet already exists: %s", cmd.Target.Sheet)
		}
	case "clear_cell", "update_cell", "insert_cells":
		if strings.TrimSpace(cmd.Target.Sheet) == "" || strings.TrimSpace(cmd.Target.Cell) == "" {
			return fmt.Errorf("%s requires target.sheet and target.cell", cmd.Op)
		}
		if r.book.engine.Book.SheetByName(cmd.Target.Sheet) == nil {
			return fmt.Errorf("找不到 sheet: %s", cmd.Target.Sheet)
		}
	default:
		if len(spec.RequiredArgs) > 0 && strings.TrimSpace(cmd.Target.Sheet) == "" {
			return fmt.Errorf("%s requires target.sheet or explicit command arguments", cmd.Op)
		}
	}
	return nil
}

func (r *interactionRun) requiresConfirmation(spec ops.OperationSpec) bool {
	if spec.RequiresConfirmation() {
		return true
	}
	levels := r.options.RequireConfirmationForLevels
	if levels == nil {
		levels = []ops.OperationLevel{ops.LevelStructureEdit, ops.LevelDestructive, ops.LevelExternalWrite}
	}
	for _, level := range levels {
		if spec.Level == level {
			return true
		}
	}
	return false
}

func (r *interactionRun) defaultSheet() string {
	if r.book == nil || r.book.engine == nil || r.book.engine.Book == nil || len(r.book.engine.Book.Sheets) != 1 {
		return ""
	}
	return r.book.engine.Book.Sheets[0].Name
}

func (r *interactionRun) emit(ctx context.Context, state InteractionState, message string, metadata map[string]any) {
	r.step++
	event := StateEvent{
		RunID:       r.runID,
		Step:        r.step,
		State:       state,
		Message:     message,
		UserRequest: r.req.UserRequest,
		Metadata:    metadata,
		CreatedAt:   time.Now(),
	}
	r.events = append(r.events, event)
	if r.options.EventSink != nil {
		_ = r.options.EventSink.OnStateEvent(ctx, event)
	}
}

func (r *interactionRun) emitWithCandidates(ctx context.Context, state InteractionState, message string, candidates []OperationCandidate, cmd *Command, metadata map[string]any) {
	r.step++
	event := StateEvent{
		RunID:       r.runID,
		Step:        r.step,
		State:       state,
		Message:     message,
		UserRequest: r.req.UserRequest,
		Candidates:  candidates,
		Command:     cmd,
		Metadata:    metadata,
		CreatedAt:   time.Now(),
	}
	r.events = append(r.events, event)
	if r.options.EventSink != nil {
		_ = r.options.EventSink.OnStateEvent(ctx, event)
	}
}

func (r *interactionRun) emitWithCommand(ctx context.Context, state InteractionState, message string, cmd *Command, metadata map[string]any) {
	eventMeta := metadata
	spec, ok := r.registry.Get(cmd.Op)
	level, risk := "", ""
	if ok {
		level = string(spec.Level)
		risk = string(spec.Risk)
	}
	r.step++
	event := StateEvent{
		RunID:       r.runID,
		Step:        r.step,
		State:       state,
		Message:     message,
		UserRequest: r.req.UserRequest,
		Op:          cmd.Op,
		Level:       level,
		Risk:        risk,
		Command:     cmd,
		Metadata:    eventMeta,
		CreatedAt:   time.Now(),
	}
	r.events = append(r.events, event)
	if r.options.EventSink != nil {
		_ = r.options.EventSink.OnStateEvent(ctx, event)
	}
}

func (r *interactionRun) emitWithDiff(ctx context.Context, state InteractionState, message string, cmd *Command, diff *Diff) {
	r.step++
	event := StateEvent{
		RunID:       r.runID,
		Step:        r.step,
		State:       state,
		Message:     message,
		UserRequest: r.req.UserRequest,
		Op:          cmd.Op,
		Command:     cmd,
		Diff:        diff,
		CreatedAt:   time.Now(),
	}
	r.events = append(r.events, event)
	if r.options.EventSink != nil {
		_ = r.options.EventSink.OnStateEvent(ctx, event)
	}
}

func (r *interactionRun) fail(ctx context.Context, cmd *Command, err error) (InteractionResult, error) {
	r.step++
	event := StateEvent{
		RunID:       r.runID,
		Step:        r.step,
		State:       StateFailed,
		Message:     "interaction failed",
		UserRequest: r.req.UserRequest,
		Command:     cmd,
		Error:       err.Error(),
		CreatedAt:   time.Now(),
	}
	r.events = append(r.events, event)
	if r.options.EventSink != nil {
		_ = r.options.EventSink.OnStateEvent(ctx, event)
	}
	return r.result(StateFailed, err.Error(), cmd, nil, OperationRecord{}, nil), err
}

func (r *interactionRun) result(status InteractionState, message string, cmd *Command, diff *Diff, record OperationRecord, candidates []OperationCandidate) InteractionResult {
	return InteractionResult{
		RunID:      r.runID,
		Status:     status,
		Message:    message,
		Command:    cmd,
		Diff:       diff,
		Record:     record,
		Candidates: candidates,
		Events:     append([]StateEvent(nil), r.events...),
	}
}

func fallbackCandidates(query string, registry *ops.Registry) []OperationCandidate {
	if registry == nil {
		return nil
	}
	tokens := fallbackTokens(query)
	var out []OperationCandidate
	for _, spec := range registry.List() {
		text := strings.ToLower(strings.Join(append([]string{spec.Op, spec.Description}, spec.UserPhrases...), " "))
		score := 0.0
		for _, token := range tokens {
			if strings.Contains(text, token) {
				score += float64(len(token))
			}
		}
		if score > 0 {
			out = append(out, OperationCandidate{
				Op:    spec.Op,
				Score: score,
				Kind:  "op_spec",
				Level: string(spec.Level),
				Risk:  string(spec.Risk),
			})
		}
	}
	return out
}

func fallbackTokens(text string) []string {
	text = strings.TrimSpace(strings.ToLower(text))
	if text == "" {
		return nil
	}
	tokens := strings.Fields(text)
	if len(tokens) == 0 {
		tokens = append(tokens, text)
	}
	if len(tokens) == 1 && tokens[0] != text {
		tokens = append(tokens, text)
	}
	for _, r := range text {
		if r > 127 {
			tokens = append(tokens, string(r))
		}
	}
	seen := make(map[string]bool)
	out := make([]string, 0, len(tokens))
	for _, token := range tokens {
		if token == "" || seen[token] {
			continue
		}
		seen[token] = true
		out = append(out, token)
	}
	return out
}

func confirmationMessage(cmd Command, spec ops.OperationSpec) string {
	switch cmd.Op {
	case "create_sheet":
		return fmt.Sprintf("将创建一个名为“%s”的新工作表，是否继续？", cmd.Target.Sheet)
	case "delete_sheet":
		return fmt.Sprintf("将删除工作表“%s”，此操作风险为 %s，是否继续？", cmd.Target.Sheet, spec.Risk)
	default:
		return fmt.Sprintf("将执行 %s，风险为 %s，是否继续？", cmd.Op, spec.Risk)
	}
}

var cellAddressPattern = regexp.MustCompile(`(?i)\b[A-Z]{1,3}[1-9][0-9]*\b`)

func extractCellAddress(text string) string {
	match := cellAddressPattern.FindString(text)
	return strings.ToUpper(match)
}

func extractInsertShift(text string) string {
	lower := strings.ToLower(text)
	if strings.Contains(lower, "down") || strings.Contains(text, "下移") || strings.Contains(text, "向下") {
		return "down"
	}
	return "right"
}

func extractSheetName(text string) string {
	cleaned := strings.TrimSpace(text)
	cleaned = strings.Trim(cleaned, "。.!！?？\"'“”‘’")
	replacers := []string{
		"帮我", "", "请", "", "新建", "", "创建", "", "新增", "",
		"加一个", "", "一个", "", "名为", "", "叫做", "",
		"sheet", "", "Sheet", "", "工作表", "", "表格", "", "表", "", "页", "",
	}
	replacer := strings.NewReplacer(replacers...)
	cleaned = strings.TrimSpace(replacer.Replace(cleaned))
	cleaned = strings.Trim(cleaned, "。.!！?？\"'“”‘’")
	return cleaned
}
