package memory

import (
	"context"
	"sort"
	"strings"
	"sync"
	"time"
	"unicode"

	"github.com/iEvan-lhr/go-excel-agent/ops"
)

type VectorRecordKind string

const (
	VectorKindOperationSpec     VectorRecordKind = "op_spec"
	VectorKindOperationExample  VectorRecordKind = "op_example"
	VectorKindSuccessExperience VectorRecordKind = "success_experience"
	VectorKindFailureCase       VectorRecordKind = "failure_case"
	VectorKindUserPreference    VectorRecordKind = "user_preference"
)

type VectorRecord struct {
	ID        string           `json:"id"`
	Kind      VectorRecordKind `json:"kind"`
	Text      string           `json:"text"`
	Embedding []float32        `json:"embedding,omitempty"`

	Op    string             `json:"op,omitempty"`
	Level ops.OperationLevel `json:"level,omitempty"`
	Risk  ops.RiskLevel      `json:"risk,omitempty"`
	Tags  []string           `json:"tags,omitempty"`

	ArgsPattern map[string]string  `json:"args_pattern,omitempty"`
	ActualArgs  map[string]any     `json:"actual_args,omitempty"`
	Constraints map[string]any     `json:"constraints,omitempty"`
	ScoreHints  map[string]float64 `json:"score_hints,omitempty"`

	WorkbookID string   `json:"workbook_id,omitempty"`
	Sheet      string   `json:"sheet,omitempty"`
	Columns    []string `json:"columns,omitempty"`

	SuccessCount int `json:"success_count,omitempty"`
	FailureCount int `json:"failure_count,omitempty"`

	SourceOperationID string    `json:"source_operation_id,omitempty"`
	CreatedAt         time.Time `json:"created_at"`
	LastUsedAt        time.Time `json:"last_used_at,omitempty"`
}

type VectorSearchRequest struct {
	Query      string             `json:"query"`
	Kinds      []VectorRecordKind `json:"kinds,omitempty"`
	Ops        []string           `json:"ops,omitempty"`
	WorkbookID string             `json:"workbook_id,omitempty"`
	Limit      int                `json:"limit,omitempty"`
}

type VectorSearchResult struct {
	Record VectorRecord `json:"record"`
	Score  float64      `json:"score"`
	Reason string       `json:"reason,omitempty"`
}

type VectorStore interface {
	Upsert(ctx context.Context, records []VectorRecord) error
	Search(ctx context.Context, req VectorSearchRequest) ([]VectorSearchResult, error)
}

type InMemoryVectorStore struct {
	mu      sync.RWMutex
	records map[string]VectorRecord
}

func NewInMemoryVectorStore() *InMemoryVectorStore {
	return &InMemoryVectorStore{records: make(map[string]VectorRecord)}
}

func (s *InMemoryVectorStore) Upsert(ctx context.Context, records []VectorRecord) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.records == nil {
		s.records = make(map[string]VectorRecord)
	}
	now := time.Now()
	for _, record := range records {
		if record.ID == "" {
			continue
		}
		if record.CreatedAt.IsZero() {
			record.CreatedAt = now
		}
		s.records[record.ID] = record
	}
	return nil
}

func (s *InMemoryVectorStore) Search(ctx context.Context, req VectorSearchRequest) ([]VectorSearchResult, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	limit := req.Limit
	if limit <= 0 {
		limit = 8
	}
	kinds := makeSet(req.Kinds)
	opsFilter := makeStringSet(req.Ops)
	tokens := vectorTokens(req.Query)

	s.mu.RLock()
	defer s.mu.RUnlock()
	results := make([]VectorSearchResult, 0, len(s.records))
	for _, record := range s.records {
		if len(kinds) > 0 && !kinds[record.Kind] {
			continue
		}
		if len(opsFilter) > 0 && !opsFilter[strings.ToLower(record.Op)] {
			continue
		}
		if req.WorkbookID != "" && record.WorkbookID != "" && record.WorkbookID != req.WorkbookID {
			continue
		}
		score := scoreVectorRecord(record, tokens)
		if score <= 0 {
			continue
		}
		results = append(results, VectorSearchResult{Record: record, Score: score, Reason: "lexical_match"})
	}
	sort.Slice(results, func(i, j int) bool {
		if results[i].Score == results[j].Score {
			return results[i].Record.ID < results[j].Record.ID
		}
		return results[i].Score > results[j].Score
	})
	if len(results) > limit {
		results = results[:limit]
	}
	return results, nil
}

func SeedOperationSpecs(ctx context.Context, store VectorStore, specs []ops.OperationSpec) error {
	if store == nil {
		return nil
	}
	records := make([]VectorRecord, 0, len(specs))
	for _, spec := range specs {
		text := strings.Join(append([]string{spec.Op, spec.Description}, spec.UserPhrases...), " ")
		records = append(records, VectorRecord{
			ID:        "op_spec:" + spec.Op + ":" + spec.Version,
			Kind:      VectorKindOperationSpec,
			Text:      text,
			Op:        spec.Op,
			Level:     spec.Level,
			Risk:      spec.Risk,
			Tags:      append([]string(nil), spec.Tags...),
			CreatedAt: time.Now(),
		})
	}
	return store.Upsert(ctx, records)
}

func makeSet(values []VectorRecordKind) map[VectorRecordKind]bool {
	out := make(map[VectorRecordKind]bool, len(values))
	for _, value := range values {
		out[value] = true
	}
	return out
}

func makeStringSet(values []string) map[string]bool {
	out := make(map[string]bool, len(values))
	for _, value := range values {
		value = strings.ToLower(strings.TrimSpace(value))
		if value != "" {
			out[value] = true
		}
	}
	return out
}

func scoreVectorRecord(record VectorRecord, tokens []string) float64 {
	if len(tokens) == 0 {
		return 0
	}
	text := strings.ToLower(strings.Join(append([]string{record.Text, record.Op}, record.Tags...), " "))
	var score float64
	for _, token := range tokens {
		if strings.Contains(text, token) {
			score += float64(len(token))
		}
	}
	if record.SuccessCount > 0 {
		score += float64(record.SuccessCount) * 0.25
	}
	if record.FailureCount > 0 {
		score -= float64(record.FailureCount) * 0.5
	}
	return score
}

func vectorTokens(text string) []string {
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
	for _, r := range text {
		if r > unicode.MaxASCII && (unicode.IsLetter(r) || unicode.IsNumber(r)) {
			tokens = append(tokens, string(r))
		}
	}
	return uniqueVectorTokens(tokens)
}

func uniqueVectorTokens(values []string) []string {
	seen := make(map[string]bool)
	out := make([]string, 0, len(values))
	for _, value := range values {
		if value == "" || seen[value] {
			continue
		}
		seen[value] = true
		out = append(out, value)
	}
	return out
}
