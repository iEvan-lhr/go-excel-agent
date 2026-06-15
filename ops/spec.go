package ops

type OperationLevel string

const (
	LevelReadOnly      OperationLevel = "read_only"
	LevelLocate        OperationLevel = "locate"
	LevelCellEdit      OperationLevel = "cell_edit"
	LevelRangeEdit     OperationLevel = "range_edit"
	LevelStructureEdit OperationLevel = "structure_edit"
	LevelDestructive   OperationLevel = "destructive"
	LevelExternalWrite OperationLevel = "external_write"
)

type RiskLevel string

const (
	RiskLow    RiskLevel = "low"
	RiskMedium RiskLevel = "medium"
	RiskHigh   RiskLevel = "high"
)

type ConfirmationPolicy string

const (
	ConfirmNever     ConfirmationPolicy = "never"
	ConfirmSometimes ConfirmationPolicy = "sometimes"
	ConfirmAlways    ConfirmationPolicy = "always"
)

type ArgSpec struct {
	Name        string `json:"name"`
	Type        string `json:"type"`
	Required    bool   `json:"required"`
	Description string `json:"description,omitempty"`
}

type OperationExample struct {
	UserRequest string         `json:"user_request"`
	Command     map[string]any `json:"command"`
}

type OperationSpec struct {
	Op                 string             `json:"op"`
	Version            string             `json:"version"`
	Level              OperationLevel     `json:"level"`
	Risk               RiskLevel          `json:"risk"`
	ConfirmationPolicy ConfirmationPolicy `json:"confirmation_policy"`
	Description        string             `json:"description"`
	UserPhrases        []string           `json:"user_phrases,omitempty"`
	Tags               []string           `json:"tags,omitempty"`
	RequiredArgs       []string           `json:"required_args,omitempty"`
	OptionalArgs       []string           `json:"optional_args,omitempty"`
	ArgSchema          map[string]ArgSpec `json:"arg_schema,omitempty"`
	Preconditions      []string           `json:"preconditions,omitempty"`
	Postconditions     []string           `json:"postconditions,omitempty"`
	CanUndo            bool               `json:"can_undo"`
	Examples           []OperationExample `json:"examples,omitempty"`
}

func (s OperationSpec) RequiresConfirmation() bool {
	return s.ConfirmationPolicy == ConfirmAlways ||
		s.Risk == RiskHigh ||
		s.Level == LevelDestructive
}
