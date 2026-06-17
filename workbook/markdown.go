package workbook

import (
	"strings"
)

// ToMarkdown converts a Sheet to a Markdown table string in GFM format.
func (s *Sheet) ToMarkdown() string {
	var sb strings.Builder
	sb.WriteString("# " + s.Name + "\n\n")

	if len(s.Rows) == 0 {
		sb.WriteString("Empty sheet.\n")
		return sb.String()
	}

	maxCols := 0
	for _, r := range s.Rows {
		if len(r) > maxCols {
			maxCols = len(r)
		}
	}

	if maxCols == 0 {
		sb.WriteString("Empty sheet.\n")
		return sb.String()
	}

	escapeCell := func(val string) string {
		val = strings.ReplaceAll(val, "\\", "\\\\")
		val = strings.ReplaceAll(val, "|", "\\|")
		val = strings.ReplaceAll(val, "\r\n", "<br>")
		val = strings.ReplaceAll(val, "\n", "<br>")
		return val
	}

	// First row as header
	headerRow := s.Rows[0]
	sb.WriteString("|")
	for i := 0; i < maxCols; i++ {
		val := ""
		if i < len(headerRow) {
			val = escapeCell(headerRow[i])
		}
		sb.WriteString(" " + val + " |")
	}
	sb.WriteString("\n")

	// Separator row
	sb.WriteString("|")
	for i := 0; i < maxCols; i++ {
		sb.WriteString(" --- |")
	}
	sb.WriteString("\n")

	// Data rows
	for idx := 1; idx < len(s.Rows); idx++ {
		row := s.Rows[idx]
		sb.WriteString("|")
		for i := 0; i < maxCols; i++ {
			val := ""
			if i < len(row) {
				val = escapeCell(row[i])
			}
			sb.WriteString(" " + val + " |")
		}
		sb.WriteString("\n")
	}

	return sb.String()
}
