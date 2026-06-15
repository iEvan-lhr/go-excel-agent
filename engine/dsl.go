package engine

type Command struct {
	Op     string `json:"op"`
	Target Target `json:"target,omitempty"`
	Args   any    `json:"args,omitempty"`
}

type Target struct {
	Sheet        string `json:"sheet,omitempty"`
	Cell         string `json:"cell,omitempty"`
	Range        string `json:"range,omitempty"`
	Column       string `json:"column,omitempty"`
	SearchQuery  string `json:"search_query,omitempty"`
	SearchColumn string `json:"search_column,omitempty"`
	Scope        *Scope `json:"scope,omitempty"`
}

type Scope struct {
	Type         string `json:"type"`
	Range        string `json:"range,omitempty"`
	Query        string `json:"query,omitempty"`
	SearchColumn string `json:"search_column,omitempty"`
}

type FindRequest struct {
	Sheet        string `json:"sheet,omitempty"`
	Type         string `json:"type,omitempty"`
	Range        string `json:"range,omitempty"`
	Query        string `json:"query,omitempty"`
	SearchColumn string `json:"search_column,omitempty"`
}

type UpdateCellRequest struct {
	Sheet string `json:"sheet,omitempty"`
	Cell  string `json:"cell,omitempty"`
	Value any    `json:"value,omitempty"`
}

type ClearCellRequest struct {
	Sheet string `json:"sheet,omitempty"`
	Cell  string `json:"cell,omitempty"`
}

type CreateSheetRequest struct {
	Sheet      string `json:"sheet,omitempty"`
	AfterSheet string `json:"after_sheet,omitempty"`
}

type InsertCellsRequest struct {
	Sheet string `json:"sheet,omitempty"`
	Cell  string `json:"cell,omitempty"`
	Shift string `json:"shift,omitempty"`
}

type BatchUpdateRequest struct {
	Sheet        string       `json:"sheet,omitempty"`
	Scope        Scope        `json:"scope"`
	TargetColumn string       `json:"target_column,omitempty"`
	Action       UpdateAction `json:"action"`
}

type UpdateAction struct {
	Type    string `json:"type,omitempty"`
	Value   any    `json:"value,omitempty"`
	Find    string `json:"find,omitempty"`
	Replace string `json:"replace,omitempty"`
}

type AggregateRequest struct {
	Sheet  string `json:"sheet,omitempty"`
	Column string `json:"column,omitempty"`
	Type   string `json:"type,omitempty"`
	Scope  *Scope `json:"scope,omitempty"`
}

type FindArgs struct {
	Type         string `json:"type,omitempty"`
	Range        string `json:"range,omitempty"`
	Query        string `json:"query,omitempty"`
	SearchColumn string `json:"search_column,omitempty"`
}

type UpdateCellArgs struct {
	Value any `json:"value,omitempty"`
}

type ClearCellArgs struct{}

type CreateSheetArgs struct {
	AfterSheet string `json:"after_sheet,omitempty"`
}

type InsertCellsArgs struct {
	Shift string `json:"shift,omitempty"`
}

type BatchUpdateArgs struct {
	Action  string `json:"action,omitempty"`
	Value   any    `json:"value,omitempty"`
	Find    string `json:"find,omitempty"`
	Replace string `json:"replace,omitempty"`
}

type AggregateArgs struct {
	Column string `json:"column,omitempty"`
	Type   string `json:"type,omitempty"`
}
