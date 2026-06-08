package memory

import (
	"context"
	"encoding/json"
	"strings"
)

type ModelMessage struct {
	Role    string `json:"role"`
	Content string `json:"content"`
}

type ModelRequest struct {
	SystemPrompt string         `json:"system_prompt,omitempty"`
	Messages     []ModelMessage `json:"messages,omitempty"`
	Prompt       string         `json:"prompt,omitempty"`
	Metadata     map[string]any `json:"metadata,omitempty"`
}

type ModelResponse struct {
	Content  string         `json:"content"`
	Metadata map[string]any `json:"metadata,omitempty"`
}

type StreamCallback func(ctx context.Context, chunk string) error

type TextModel interface {
	Complete(ctx context.Context, req ModelRequest) (ModelResponse, error)
}

type StreamingTextModel interface {
	TextModel
	CompleteStream(ctx context.Context, req ModelRequest, cb StreamCallback) (ModelResponse, error)
}

type TextModelFunc func(ctx context.Context, req ModelRequest) (ModelResponse, error)

func (f TextModelFunc) Complete(ctx context.Context, req ModelRequest) (ModelResponse, error) {
	return f(ctx, req)
}

type ColumnTagRequest struct {
	WorkbookID   string         `json:"workbook_id,omitempty"`
	SheetName    string         `json:"sheet_name,omitempty"`
	ColumnName   string         `json:"column_name"`
	ColumnIndex  int            `json:"column_index"`
	ColumnLetter string         `json:"column_letter"`
	InferredType string         `json:"inferred_type"`
	SampleValues []string       `json:"sample_values,omitempty"`
	ValueHints   []ValuePattern `json:"value_hints,omitempty"`
}

type ColumnTagger interface {
	Tags(ctx context.Context, req ColumnTagRequest) ([]string, error)
}

type ColumnTaggerFunc func(ctx context.Context, req ColumnTagRequest) ([]string, error)

func (f ColumnTaggerFunc) Tags(ctx context.Context, req ColumnTagRequest) ([]string, error) {
	return f(ctx, req)
}

type IntentGeneralizer interface {
	GeneralizeOperation(ctx context.Context, record OperationRecord) (GeneralizedIntent, error)
}

type IntentGeneralizerFunc func(ctx context.Context, record OperationRecord) (GeneralizedIntent, error)

func (f IntentGeneralizerFunc) GeneralizeOperation(ctx context.Context, record OperationRecord) (GeneralizedIntent, error) {
	return f(ctx, record)
}

type ExecutionSummarizer interface {
	SummarizeOperation(ctx context.Context, record OperationRecord) (ExecutionSummary, error)
}

type ExecutionSummarizerFunc func(ctx context.Context, record OperationRecord) (ExecutionSummary, error)

func (f ExecutionSummarizerFunc) SummarizeOperation(ctx context.Context, record OperationRecord) (ExecutionSummary, error) {
	return f(ctx, record)
}

type ModelColumnTagger struct {
	Model       TextModel
	BuildPrompt func(ColumnTagRequest) ModelRequest
	Parse       func(ModelResponse) ([]string, error)
}

func (t ModelColumnTagger) Tags(ctx context.Context, req ColumnTagRequest) ([]string, error) {
	if t.Model == nil {
		return nil, nil
	}
	build := t.BuildPrompt
	if build == nil {
		build = defaultColumnTagPrompt
	}
	parse := t.Parse
	if parse == nil {
		parse = parseJSONTags
	}
	resp, err := t.Model.Complete(ctx, build(req))
	if err != nil {
		return nil, err
	}
	return parse(resp)
}

type RuleBasedColumnTagger struct {
	Rules []SemanticTagRule
}

type SemanticTagRule struct {
	Tag   string   `json:"tag"`
	Terms []string `json:"terms"`
}

func (t RuleBasedColumnTagger) Tags(ctx context.Context, req ColumnTagRequest) ([]string, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	lower := strings.ToLower(strings.TrimSpace(req.ColumnName))
	var tags []string
	for _, rule := range t.Rules {
		for _, term := range rule.Terms {
			if strings.Contains(lower, strings.ToLower(strings.TrimSpace(term))) {
				tags = append(tags, rule.Tag)
				break
			}
		}
	}
	return uniqueStrings(tags), nil
}

func defaultColumnTagPrompt(req ColumnTagRequest) ModelRequest {
	raw, _ := json.Marshal(req)
	return ModelRequest{
		SystemPrompt: "Return a compact JSON array of semantic tags for the spreadsheet column. Do not include explanations.",
		Prompt:       string(raw),
		Metadata: map[string]any{
			"task": "column_semantic_tags",
		},
	}
}

func parseJSONTags(resp ModelResponse) ([]string, error) {
	var tags []string
	if err := json.Unmarshal([]byte(strings.TrimSpace(resp.Content)), &tags); err != nil {
		return nil, err
	}
	return uniqueStrings(tags), nil
}
