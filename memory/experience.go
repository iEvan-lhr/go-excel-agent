package memory

import (
	"fmt"
	"strings"
	"time"

	"github.com/iEvan-lhr/go-excel-agent/ops"
)

func VectorRecordFromOperation(record OperationRecord, spec ops.OperationSpec) VectorRecord {
	kind := VectorKindSuccessExperience
	if record.Error != "" {
		kind = VectorKindFailureCase
	}
	id := record.OperationID
	if id == "" {
		id = fmt.Sprintf("%d", time.Now().UnixNano())
	}
	text := buildExperienceText(record)
	out := VectorRecord{
		ID:                "experience:" + id,
		Kind:              kind,
		Text:              text,
		Op:                strings.ToLower(strings.TrimSpace(record.Command.Op)),
		Level:             spec.Level,
		Risk:              spec.Risk,
		Tags:              append([]string(nil), spec.Tags...),
		ArgsPattern:       inferArgsPattern(record),
		ActualArgs:        actualArgs(record),
		WorkbookID:        record.WorkbookID,
		Sheet:             record.Command.Target.Sheet,
		SourceOperationID: record.OperationID,
		CreatedAt:         time.Now(),
	}
	if record.Error == "" {
		out.SuccessCount = 1
	} else {
		out.FailureCount = 1
		out.Constraints = map[string]any{"error": record.Error}
	}
	return out
}

func buildExperienceText(record OperationRecord) string {
	parts := []string{}
	if record.UserRequest != "" {
		parts = append(parts, "user request: "+record.UserRequest)
	}
	if record.Command.Op != "" {
		parts = append(parts, "operation: "+record.Command.Op)
	}
	if record.Command.Target.Sheet != "" {
		parts = append(parts, "sheet: "+record.Command.Target.Sheet)
	}
	if record.Command.Target.Cell != "" {
		parts = append(parts, "cell: "+record.Command.Target.Cell)
	}
	if record.Command.Target.Column != "" {
		parts = append(parts, "column: "+record.Command.Target.Column)
	}
	if len(parts) == 0 {
		return string(record.CommandJSON)
	}
	return strings.Join(parts, "\n")
}

func inferArgsPattern(record OperationRecord) map[string]string {
	pattern := make(map[string]string)
	if record.Command.Target.Sheet != "" {
		pattern["sheet"] = "<sheet_name>"
	}
	if record.Command.Target.Cell != "" {
		pattern["cell"] = "<cell_address>"
	}
	if record.Command.Target.Range != "" {
		pattern["range"] = "<range>"
	}
	if record.Command.Target.Column != "" {
		pattern["column"] = "<column_name>"
	}
	if len(pattern) == 0 {
		return nil
	}
	return pattern
}

func actualArgs(record OperationRecord) map[string]any {
	args := make(map[string]any)
	if record.Command.Target.Sheet != "" {
		args["sheet"] = record.Command.Target.Sheet
	}
	if record.Command.Target.Cell != "" {
		args["cell"] = record.Command.Target.Cell
	}
	if record.Command.Target.Range != "" {
		args["range"] = record.Command.Target.Range
	}
	if record.Command.Target.Column != "" {
		args["column"] = record.Command.Target.Column
	}
	if record.Command.Args != nil {
		args["args"] = record.Command.Args
	}
	if len(args) == 0 {
		return nil
	}
	return args
}
