package memory

import (
	"context"
	"encoding/json"
	"fmt"
	"sort"
	"strings"
	"time"
	"unicode"

	"github.com/iEvan-lhr/go-excel-agent/engine"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
)

type Store struct {
	Artifacts  map[string]ArtifactMemory
	Graphs     map[string]DataGraph
	Operations []OperationRecord
	Summaries  []ExecutionSummary
	Session    SessionFocus
	Options    Options

	nextOperation int
	nextSummary   int
	nextCapsule   int
}

func NewStore(options ...Option) *Store {
	return &Store{
		Artifacts: make(map[string]ArtifactMemory),
		Graphs:    make(map[string]DataGraph),
		Options:   applyOptions(options),
	}
}

func (s *Store) IndexWorkbook(workbookID string, book *workbook.Workbook) ArtifactMemory {
	artifact, _ := s.IndexWorkbookContext(context.Background(), workbookID, book)
	return artifact
}

func (s *Store) IndexWorkbookContext(ctx context.Context, workbookID string, book *workbook.Workbook) (ArtifactMemory, error) {
	s.ensure()
	artifact, err := BuildArtifactMemoryWithOptions(ctx, workbookID, book, s.Options)
	if err != nil {
		return ArtifactMemory{}, err
	}
	graph := BuildDataGraph(artifact)
	s.Artifacts[artifact.WorkbookID] = artifact
	s.Graphs[artifact.WorkbookID] = graph
	s.Session.WorkbookID = artifact.WorkbookID
	s.Session.UpdatedAt = time.Now()
	return artifact, nil
}

func (s *Store) RecordOperation(record OperationRecord) (OperationRecord, ExecutionSummary, error) {
	return s.RecordOperationContext(context.Background(), record)
}

func (s *Store) RecordOperationContext(ctx context.Context, record OperationRecord) (OperationRecord, ExecutionSummary, error) {
	s.ensure()
	if record.WorkbookID == "" {
		record.WorkbookID = s.Session.WorkbookID
	}
	if record.WorkbookID == "" {
		return record, ExecutionSummary{}, fmt.Errorf("operation workbook_id is required")
	}
	if record.OperationID == "" {
		s.nextOperation++
		record.OperationID = fmt.Sprintf("op_%06d", s.nextOperation)
	}
	if record.Timestamp.IsZero() {
		record.Timestamp = time.Now()
	}
	if len(record.CommandJSON) == 0 {
		raw, err := json.Marshal(record.Command)
		if err != nil {
			return record, ExecutionSummary{}, fmt.Errorf("marshal command: %w", err)
		}
		record.CommandJSON = raw
	}
	if record.GeneralizedIntent.IntentType == "" {
		intent, err := s.generalizeOperation(ctx, record)
		if err != nil {
			return record, ExecutionSummary{}, err
		}
		record.GeneralizedIntent = intent
	}
	if len(record.Locations) == 0 && record.Diff != nil {
		record.Locations = locationsFromDiff(record.WorkbookID, record.Diff)
	}

	s.Operations = append(s.Operations, record)
	summary, err := s.summarizeOperation(ctx, record)
	if err != nil {
		return record, ExecutionSummary{}, err
	}
	s.Summaries = append(s.Summaries, summary)
	s.updateSessionFocus(record)
	s.updateGraphWithOperation(record, summary)
	return record, summary, nil
}

func (s *Store) BuildContextCapsule(book *workbook.Workbook, req ContextRequest) (ContextCapsule, error) {
	return s.BuildContextCapsuleContext(context.Background(), book, req)
}

func (s *Store) BuildContextCapsuleContext(ctx context.Context, book *workbook.Workbook, req ContextRequest) (ContextCapsule, error) {
	s.ensure()
	if req.WorkbookID == "" {
		req.WorkbookID = s.Session.WorkbookID
	}
	if req.WorkbookID == "" {
		req.WorkbookID = "current"
	}
	if _, ok := s.Artifacts[req.WorkbookID]; !ok && book != nil {
		if _, err := s.IndexWorkbookContext(ctx, req.WorkbookID, book); err != nil {
			return ContextCapsule{}, err
		}
	}
	artifact, ok := s.Artifacts[req.WorkbookID]
	if !ok {
		return ContextCapsule{}, fmt.Errorf("workbook memory not found: %s", req.WorkbookID)
	}
	graph := s.Graphs[req.WorkbookID]

	budget := normalizeBudget(req)
	nodes := selectGraphNodes(graph, req.Query, budget.MaxNodes)
	rows := retrieveRows(artifact, book, req, budget.MaxRows, budget.MaxCells)
	ops, summaries := s.selectOperationContext(req)

	s.nextCapsule++
	capsule := ContextCapsule{
		CapsuleID:          fmt.Sprintf("ctx_%06d", s.nextCapsule),
		WorkbookID:         req.WorkbookID,
		Purpose:            req.Purpose,
		Query:              req.Query,
		IncludedNodes:      nodes,
		EvidenceRows:       rows,
		RelevantOperations: ops,
		ExecutionSummaries: summaries,
		Budget:             budget,
		CreatedAt:          time.Now(),
	}
	capsule.Budget.IncludedRows = len(rows)
	capsule.Budget.IncludedNodes = len(nodes)
	for _, row := range rows {
		capsule.Budget.IncludedCells += len(row.Cells)
	}
	if len(rows) == budget.MaxRows {
		capsule.Excluded = append(capsule.Excluded, "raw rows beyond evidence budget")
	}
	if len(nodes) == budget.MaxNodes {
		capsule.Excluded = append(capsule.Excluded, "graph nodes beyond evidence budget")
	}

	s.Session.WorkbookID = req.WorkbookID
	s.Session.LastContextCapsuleID = capsule.CapsuleID
	s.Session.LastRelevantRows = evidenceRefs(req.WorkbookID, rows)
	s.Session.UpdatedAt = time.Now()
	return capsule, nil
}

func GeneralizeCommand(cmd engine.Command) GeneralizedIntent {
	intent := GeneralizedIntent{Properties: make(map[string]any)}
	switch strings.ToLower(strings.TrimSpace(cmd.Op)) {
	case "find":
		intent.IntentType = "retrieve"
		intent.Locator = "search_or_range"
		intent.Target = "cells_or_rows"
	case "get_range":
		intent.IntentType = "retrieve"
		intent.Locator = "range"
		intent.Target = "range"
	case "update_cell":
		intent.IntentType = "single_cell_update"
		intent.Locator = "cell_address"
		intent.Action = "overwrite"
		intent.Target = "cell"
	case "clear_cell":
		intent.IntentType = "single_cell_update"
		intent.Locator = "cell_address"
		intent.Action = "clear"
		intent.Target = "cell"
	case "create_sheet":
		intent.IntentType = "structure_edit"
		intent.Locator = "sheet_name"
		intent.Action = "create"
		intent.Target = "sheet"
	case "insert_cells":
		intent.IntentType = "structure_edit"
		intent.Locator = "cell_address"
		intent.Action = "insert"
		intent.Target = "cell"
	case "batch_update":
		intent.IntentType = "locate_and_update"
		intent.Locator = "scope"
		intent.Action = "batch_update"
		intent.Target = "column_or_range"
	case "aggregate":
		intent.IntentType = "aggregate"
		intent.Locator = "column"
		intent.Action = "calculate"
		intent.Target = "numeric_column"
	case "inspect_workbook":
		intent.IntentType = "inspect_workbook"
		intent.Target = "workbook_structure"
	default:
		intent.IntentType = "unknown"
	}
	if cmd.Target.Sheet != "" {
		intent.Properties["sheet"] = cmd.Target.Sheet
	}
	if cmd.Target.Column != "" {
		intent.Properties["column"] = cmd.Target.Column
	}
	if cmd.Target.Cell != "" {
		intent.Properties["cell"] = cmd.Target.Cell
	}
	return intent
}

func (s *Store) ensure() {
	if s.Artifacts == nil {
		s.Artifacts = make(map[string]ArtifactMemory)
	}
	if s.Graphs == nil {
		s.Graphs = make(map[string]DataGraph)
	}
}

func (s *Store) generalizeOperation(ctx context.Context, record OperationRecord) (GeneralizedIntent, error) {
	if s.Options.IntentGeneralizer != nil {
		intent, err := s.Options.IntentGeneralizer.GeneralizeOperation(ctx, record)
		if err != nil {
			return GeneralizedIntent{}, err
		}
		if intent.IntentType != "" {
			return intent, nil
		}
	}
	return GeneralizeCommand(record.Command), nil
}

func (s *Store) summarizeOperation(ctx context.Context, record OperationRecord) (ExecutionSummary, error) {
	var summary ExecutionSummary
	if s.Options.ExecutionSummarizer != nil {
		custom, err := s.Options.ExecutionSummarizer.SummarizeOperation(ctx, record)
		if err != nil {
			return ExecutionSummary{}, err
		}
		summary = custom
	} else {
		summary = s.buildExecutionSummary(record)
	}
	if summary.SummaryID == "" {
		s.nextSummary++
		summary.SummaryID = fmt.Sprintf("sum_%06d", s.nextSummary)
	}
	if summary.OperationID == "" {
		summary.OperationID = record.OperationID
	}
	if summary.WorkbookID == "" {
		summary.WorkbookID = record.WorkbookID
	}
	if summary.Kind == "" {
		summary.Kind = record.GeneralizedIntent.IntentType
	}
	if summary.CreatedAt.IsZero() {
		summary.CreatedAt = time.Now()
	}
	if len(summary.Evidence) == 0 {
		summary.Evidence = append([]EvidenceRef(nil), record.Locations...)
	}
	return summary, nil
}

func (s *Store) buildExecutionSummary(record OperationRecord) ExecutionSummary {
	summary := ExecutionSummary{
		OperationID: record.OperationID,
		WorkbookID:  record.WorkbookID,
		Kind:        record.GeneralizedIntent.IntentType,
		CreatedAt:   time.Now(),
		Details:     make(map[string]any),
		Evidence:    append([]EvidenceRef(nil), record.Locations...),
	}
	if record.Error != "" {
		summary.Kind = "failed_operation"
		summary.Text = "operation failed: " + record.Error
		summary.Facts = append(summary.Facts, MemoryFact{Key: "error", Value: record.Error})
		return summary
	}
	if record.Diff != nil && len(record.Diff.StructureChanges) > 0 {
		summary.Kind = "structure_change"
		summary.Text = fmt.Sprintf("%s changed workbook structure", record.Command.Op)
		for _, change := range record.Diff.StructureChanges {
			summary.Facts = append(summary.Facts,
				MemoryFact{Key: "structure_change", Value: change.Type},
				MemoryFact{Key: "sheet", Value: change.Sheet},
			)
		}
		return summary
	}
	if record.Diff == nil || record.Diff.ChangedCells == 0 {
		summary.Text = fmt.Sprintf("%s completed without cell changes", record.Command.Op)
		summary.Facts = append(summary.Facts, MemoryFact{Key: "changed_cells", Value: "0"})
		return summary
	}
	if record.Diff.ChangedCells == 1 && len(record.Diff.Changes) == 1 {
		change := record.Diff.Changes[0]
		summary.Kind = "single_cell_change"
		summary.Text = fmt.Sprintf("%s!%s changed from %q to %q", change.Sheet, change.Cell, change.OldValue, change.NewValue)
		summary.Facts = append(summary.Facts,
			MemoryFact{Key: "sheet", Value: change.Sheet},
			MemoryFact{Key: "cell", Value: change.Cell},
			MemoryFact{Key: "old_value", Value: change.OldValue},
			MemoryFact{Key: "new_value", Value: change.NewValue},
		)
		return summary
	}
	summary.Kind = "multi_cell_change"
	summary.Text = fmt.Sprintf("%s changed %d cells", record.Command.Op, record.Diff.ChangedCells)
	summary.Facts = append(summary.Facts, MemoryFact{Key: "changed_cells", Value: fmt.Sprintf("%d", record.Diff.ChangedCells)})
	summary.Details["changed_rows"] = record.Diff.ChangedRows
	return summary
}

func (s *Store) updateSessionFocus(record OperationRecord) {
	s.Session.WorkbookID = record.WorkbookID
	s.Session.LastOperationID = record.OperationID
	s.Session.LastChangedCells = append([]EvidenceRef(nil), record.Locations...)
	if len(record.Locations) > 0 {
		s.Session.LastSheet = record.Locations[0].Sheet
	}
	s.Session.UpdatedAt = time.Now()
}

func (s *Store) updateGraphWithOperation(record OperationRecord, summary ExecutionSummary) {
	graph := s.Graphs[record.WorkbookID]
	if graph.WorkbookID == "" {
		return
	}
	opNodeID := nodeID("operation", record.WorkbookID, record.OperationID)
	graph.Nodes = append(graph.Nodes, GraphNode{
		ID:      opNodeID,
		Type:    "operation",
		Label:   record.OperationID,
		Summary: summary.Text,
		Ref:     EvidenceRef{WorkbookID: record.WorkbookID},
		Metadata: map[string]any{
			"op":          record.Command.Op,
			"intent_type": record.GeneralizedIntent.IntentType,
		},
	})
	for _, location := range record.Locations {
		cellNodeID := nodeID("cell", record.WorkbookID, location.Sheet, location.Address)
		graph.Nodes = append(graph.Nodes, GraphNode{
			ID:    cellNodeID,
			Type:  "cell",
			Label: location.Sheet + "!" + location.Address,
			Ref:   location,
		})
		graph.Edges = append(graph.Edges, GraphEdge{From: opNodeID, To: cellNodeID, Type: "updated"})
	}
	s.Graphs[record.WorkbookID] = graph
}

func (s *Store) selectOperationContext(req ContextRequest) ([]OperationDigest, []ExecutionSummary) {
	if !req.IncludeOperations && req.Purpose != PurposeFollowup && req.Purpose != PurposeExplainResult {
		return nil, nil
	}
	tokens := tokenize(req.Query)
	var digests []OperationDigest
	for i := len(s.Operations) - 1; i >= 0; i-- {
		record := s.Operations[i]
		if req.WorkbookID != "" && record.WorkbookID != req.WorkbookID {
			continue
		}
		if len(tokens) > 0 && scoreText(record.UserRequest+" "+record.GeneralizedIntent.IntentType+" "+string(record.CommandJSON), tokens) == 0 {
			continue
		}
		changed := 0
		if record.Diff != nil {
			changed = record.Diff.ChangedCells
		}
		digests = append(digests, OperationDigest{
			OperationID:       record.OperationID,
			Timestamp:         record.Timestamp,
			UserRequest:       record.UserRequest,
			GeneralizedIntent: record.GeneralizedIntent,
			Locations:         append([]EvidenceRef(nil), record.Locations...),
			ChangedCells:      changed,
			Error:             record.Error,
		})
		if len(digests) >= 5 {
			break
		}
	}

	var summaries []ExecutionSummary
	for i := len(s.Summaries) - 1; i >= 0; i-- {
		summary := s.Summaries[i]
		if req.WorkbookID != "" && summary.WorkbookID != req.WorkbookID {
			continue
		}
		if len(tokens) > 0 && scoreText(summary.Text, tokens) == 0 {
			continue
		}
		summaries = append(summaries, summary)
		if len(summaries) >= 5 {
			break
		}
	}
	return digests, summaries
}

func selectGraphNodes(graph DataGraph, query string, maxNodes int) []GraphNode {
	if maxNodes <= 0 {
		maxNodes = 20
	}
	tokens := tokenize(query)
	type scoredNode struct {
		node  GraphNode
		score float64
	}
	var scored []scoredNode
	for _, node := range graph.Nodes {
		score := 0.1
		if node.Type == "workbook" || node.Type == "sheet" {
			score += 1
		}
		score += scoreText(node.Label+" "+node.Summary+" "+strings.Join(node.Tags, " "), tokens)
		if len(tokens) == 0 && (node.Type == "workbook" || node.Type == "sheet" || node.Type == "column") {
			score += 1
		}
		if score > 0 {
			scored = append(scored, scoredNode{node: node, score: score})
		}
	}
	sort.Slice(scored, func(i, j int) bool {
		if scored[i].score == scored[j].score {
			return scored[i].node.ID < scored[j].node.ID
		}
		return scored[i].score > scored[j].score
	})
	if len(scored) > maxNodes {
		scored = scored[:maxNodes]
	}
	nodes := make([]GraphNode, 0, len(scored))
	for _, item := range scored {
		nodes = append(nodes, item.node)
	}
	return nodes
}

func retrieveRows(artifact ArtifactMemory, book *workbook.Workbook, req ContextRequest, maxRows, maxCells int) []EvidenceRow {
	if maxRows <= 0 {
		maxRows = 20
	}
	if maxCells <= 0 {
		maxCells = 200
	}
	if book == nil {
		return representativeRows(artifact, maxRows, maxCells)
	}
	tokens := tokenize(req.Query)
	if len(tokens) == 0 && req.Purpose == PurposeUnderstandFile {
		return representativeRows(artifact, maxRows, maxCells)
	}

	type scoredRow struct {
		row EvidenceRow
	}
	var scored []scoredRow
	for _, sheet := range book.Sheets {
		profile := findSheetProfile(artifact, sheet.Name)
		headerRowIdx := 0
		if profile != nil && profile.HeaderRow > 0 {
			headerRowIdx = profile.HeaderRow - 1
		}
		headers := []string{}
		if headerRowIdx >= 0 && headerRowIdx < len(sheet.Rows) {
			headers = append(headers, sheet.Rows[headerRowIdx]...)
		}
		for rowIdx := headerRowIdx + 1; rowIdx < len(sheet.Rows); rowIdx++ {
			score, matched := scoreRow(headers, sheet.Rows[rowIdx], tokens)
			if score == 0 {
				continue
			}
			scored = append(scored, scoredRow{
				row: buildEvidenceRow(artifact.WorkbookID, sheet, rowIdx, headerRowIdx, score, matched),
			})
		}
	}
	sort.Slice(scored, func(i, j int) bool {
		if scored[i].row.Score == scored[j].row.Score {
			if scored[i].row.Sheet == scored[j].row.Sheet {
				return scored[i].row.RowIndex < scored[j].row.RowIndex
			}
			return scored[i].row.Sheet < scored[j].row.Sheet
		}
		return scored[i].row.Score > scored[j].row.Score
	})

	var rows []EvidenceRow
	cells := 0
	for _, item := range scored {
		nextCells := len(item.row.Cells)
		if len(rows) >= maxRows || cells+nextCells > maxCells {
			break
		}
		rows = append(rows, item.row)
		cells += nextCells
	}
	if len(rows) == 0 && req.Purpose == PurposeFollowup {
		return representativeRows(artifact, maxRows, maxCells)
	}
	return rows
}

func representativeRows(artifact ArtifactMemory, maxRows, maxCells int) []EvidenceRow {
	var rows []EvidenceRow
	cells := 0
	for _, sheet := range artifact.Sheets {
		for _, row := range sheet.TitleRows {
			if len(rows) >= maxRows || cells+len(row.Cells) > maxCells {
				return rows
			}
			rows = append(rows, row)
			cells += len(row.Cells)
		}
		for _, row := range sheet.SampleRows {
			if len(rows) >= maxRows || cells+len(row.Cells) > maxCells {
				return rows
			}
			rows = append(rows, row)
			cells += len(row.Cells)
		}
	}
	return rows
}

func findSheetProfile(artifact ArtifactMemory, name string) *SheetProfile {
	for i := range artifact.Sheets {
		if artifact.Sheets[i].Name == name {
			return &artifact.Sheets[i]
		}
	}
	return nil
}

func scoreRow(headers, row []string, tokens []string) (float64, []string) {
	if len(tokens) == 0 {
		return 0, nil
	}
	var score float64
	var matched []string
	for colIdx, value := range row {
		field := headerName(headers, colIdx)
		valueScore := scoreText(value, tokens)
		if valueScore > 0 {
			fieldScore := scoreText(field, tokens) * 0.25
			score += valueScore + fieldScore
			matched = append(matched, field)
		}
	}
	return score, matched
}

func locationsFromDiff(workbookID string, diff *workbook.Diff) []EvidenceRef {
	if diff == nil {
		return nil
	}
	locations := make([]EvidenceRef, 0, len(diff.Changes))
	for _, change := range diff.Changes {
		locations = append(locations, EvidenceRef{
			WorkbookID: workbookID,
			Sheet:      change.Sheet,
			Address:    change.Cell,
			Row:        change.RowIndex,
			Column:     change.ColIndex,
		})
	}
	return locations
}

func evidenceRefs(workbookID string, rows []EvidenceRow) []EvidenceRef {
	refs := make([]EvidenceRef, 0, len(rows))
	for _, row := range rows {
		refs = append(refs, EvidenceRef{
			WorkbookID: workbookID,
			Sheet:      row.Sheet,
			Address:    row.AddressRange,
			Row:        row.RowIndex,
		})
	}
	return refs
}

func normalizeBudget(req ContextRequest) ContextBudget {
	budget := ContextBudget{MaxRows: req.MaxRows, MaxCells: req.MaxCells, MaxNodes: req.MaxNodes}
	if budget.MaxRows <= 0 {
		budget.MaxRows = 20
	}
	if budget.MaxCells <= 0 {
		budget.MaxCells = 200
	}
	if budget.MaxNodes <= 0 {
		budget.MaxNodes = 20
	}
	return budget
}

func tokenize(text string) []string {
	text = strings.TrimSpace(strings.ToLower(text))
	if text == "" {
		return nil
	}
	var tokens []string
	var current strings.Builder
	flush := func() {
		if current.Len() > 0 {
			tokens = append(tokens, current.String())
			current.Reset()
		}
	}
	for _, r := range text {
		if unicode.IsLetter(r) || unicode.IsNumber(r) {
			current.WriteRune(r)
			continue
		}
		flush()
	}
	flush()
	if len(tokens) == 0 {
		return []string{text}
	}
	if len(tokens) == 1 && tokens[0] != text {
		tokens = append(tokens, text)
	}
	return uniqueStrings(tokens)
}

func scoreText(text string, tokens []string) float64 {
	if len(tokens) == 0 {
		return 0
	}
	lower := strings.ToLower(text)
	var score float64
	for _, token := range tokens {
		if token == "" {
			continue
		}
		if strings.Contains(lower, token) {
			score += float64(len(token))
		}
	}
	return score
}

func uniqueStrings(values []string) []string {
	seen := make(map[string]bool)
	var out []string
	for _, value := range values {
		if value == "" || seen[value] {
			continue
		}
		seen[value] = true
		out = append(out, value)
	}
	return out
}
