package workbook

// CellView represents a cell's display value, position, sizes, and styling details.
type CellView struct {
	Coordinate string     `json:"coordinate"` // e.g. "A1"
	Value      string     `json:"value"`      // cell display value
	Width      float64    `json:"width"`      // column width
	Height     float64    `json:"height"`     // row height
	Style      *CellStyle `json:"style,omitempty"`
	RowSpan    int        `json:"rowSpan,omitempty"`
	ColSpan    int        `json:"colSpan,omitempty"`
	IsMerged   bool       `json:"isMerged,omitempty"`
	IsOverlaid bool       `json:"isOverlaid,omitempty"`
}

// CellStyle contains visual styling properties of a cell.
type CellStyle struct {
	FontName   string  `json:"fontName,omitempty"`
	FontSize   float64 `json:"fontSize,omitempty"`
	Bold       bool    `json:"bold,omitempty"`
	Italic     bool    `json:"italic,omitempty"`
	FontColor  string  `json:"fontColor,omitempty"`  // hex color code, e.g. "FFFF0000"
	FillColor  string  `json:"fillColor,omitempty"`  // hex color code of background fill
	AlignHoriz string  `json:"alignHoriz,omitempty"` // "left", "center", "right"
	AlignVert  string  `json:"alignVert,omitempty"`  // "top", "center", "bottom"
}
