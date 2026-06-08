package memory

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/iEvan-lhr/go-excel-agent/workbook"
	"github.com/xuri/excelize/v2"
)

const defaultSampleRows = 3

func BuildArtifactMemory(workbookID string, book *workbook.Workbook) ArtifactMemory {
	artifact, _ := BuildArtifactMemoryWithOptions(context.Background(), workbookID, book, Options{})
	return artifact
}

func BuildArtifactMemoryWithOptions(ctx context.Context, workbookID string, book *workbook.Workbook, options Options) (ArtifactMemory, error) {
	if workbookID == "" {
		workbookID = "current"
	}
	if book == nil {
		return ArtifactMemory{WorkbookID: workbookID, LastIndexedAt: time.Now()}, nil
	}

	artifact := ArtifactMemory{
		WorkbookID:    workbookID,
		SourcePath:    book.SourcePath,
		Fingerprint:   fingerprintWorkbook(book),
		SheetCount:    len(book.Sheets),
		Sheets:        make([]SheetProfile, 0, len(book.Sheets)),
		LastIndexedAt: time.Now(),
	}
	for _, sheet := range book.Sheets {
		profile, err := buildSheetProfile(ctx, workbookID, sheet, options)
		if err != nil {
			return ArtifactMemory{}, err
		}
		artifact.Sheets = append(artifact.Sheets, profile)
	}
	return artifact, nil
}

func BuildDataGraph(artifact ArtifactMemory) DataGraph {
	graph := DataGraph{
		WorkbookID: artifact.WorkbookID,
		BuiltAt:    time.Now(),
	}

	workbookNode := GraphNode{
		ID:      nodeID("workbook", artifact.WorkbookID),
		Type:    "workbook",
		Label:   artifact.WorkbookID,
		Summary: fmt.Sprintf("%d sheets", artifact.SheetCount),
		Ref:     EvidenceRef{WorkbookID: artifact.WorkbookID},
	}
	graph.Nodes = append(graph.Nodes, workbookNode)

	for _, sheet := range artifact.Sheets {
		sheetID := nodeID("sheet", artifact.WorkbookID, sheet.Name)
		graph.Nodes = append(graph.Nodes, GraphNode{
			ID:      sheetID,
			Type:    "sheet",
			Label:   sheet.Name,
			Summary: fmt.Sprintf("%d rows x %d columns", sheet.RowCount, sheet.ColumnCount),
			Ref:     EvidenceRef{WorkbookID: artifact.WorkbookID, Sheet: sheet.Name},
			Metadata: map[string]any{
				"row_count":    sheet.RowCount,
				"column_count": sheet.ColumnCount,
				"header_row":   sheet.HeaderRow,
			},
		})
		graph.Edges = append(graph.Edges, GraphEdge{From: workbookNode.ID, To: sheetID, Type: "contains"})

		for _, region := range sheet.TableRegions {
			regionID := nodeID("table_region", artifact.WorkbookID, sheet.Name, region.Address)
			graph.Nodes = append(graph.Nodes, GraphNode{
				ID:      regionID,
				Type:    "table_region",
				Label:   sheet.Name + "!" + region.Address,
				Summary: fmt.Sprintf("rows %d-%d, columns %d-%d", region.StartRow, region.EndRow, region.StartCol, region.EndCol),
				Ref:     EvidenceRef{WorkbookID: artifact.WorkbookID, Sheet: sheet.Name, Address: region.Address},
			})
			graph.Edges = append(graph.Edges, GraphEdge{From: sheetID, To: regionID, Type: "has_region"})
		}

		for _, column := range sheet.Columns {
			columnID := nodeID("column", artifact.WorkbookID, sheet.Name, column.Letter)
			graph.Nodes = append(graph.Nodes, GraphNode{
				ID:      columnID,
				Type:    "column",
				Label:   sheet.Name + "." + column.Name,
				Summary: columnSummary(column),
				Tags:    append([]string(nil), column.SemanticTags...),
				Ref: EvidenceRef{
					WorkbookID: artifact.WorkbookID,
					Sheet:      sheet.Name,
					Column:     column.Index,
				},
				Metadata: map[string]any{
					"name":            column.Name,
					"letter":          column.Letter,
					"inferred_type":   column.InferredType,
					"non_empty_count": column.NonEmptyCount,
					"empty_ratio":     column.EmptyRatio,
				},
			})
			graph.Edges = append(graph.Edges, GraphEdge{From: sheetID, To: columnID, Type: "has_column"})
		}

		if len(sheet.SampleRows) > 0 {
			clusterID := nodeID("row_cluster", artifact.WorkbookID, sheet.Name, "sample")
			graph.Nodes = append(graph.Nodes, GraphNode{
				ID:      clusterID,
				Type:    "row_cluster",
				Label:   sheet.Name + ".representative_rows",
				Summary: fmt.Sprintf("%d representative rows", len(sheet.SampleRows)),
				Ref:     EvidenceRef{WorkbookID: artifact.WorkbookID, Sheet: sheet.Name},
				Metadata: map[string]any{
					"strategy": "structure_sample",
					"rows":     rowIndexes(sheet.SampleRows),
				},
			})
			graph.Edges = append(graph.Edges, GraphEdge{From: sheetID, To: clusterID, Type: "has_sample"})
		}

		if len(sheet.TitleRows) > 0 {
			titleID := nodeID("row_cluster", artifact.WorkbookID, sheet.Name, "title")
			graph.Nodes = append(graph.Nodes, GraphNode{
				ID:      titleID,
				Type:    "row_cluster",
				Label:   sheet.Name + ".title_rows",
				Summary: fmt.Sprintf("%d pre-header title/description rows", len(sheet.TitleRows)),
				Ref:     EvidenceRef{WorkbookID: artifact.WorkbookID, Sheet: sheet.Name},
				Metadata: map[string]any{
					"strategy": "pre_header_evidence",
					"rows":     rowIndexes(sheet.TitleRows),
				},
			})
			graph.Edges = append(graph.Edges, GraphEdge{From: sheetID, To: titleID, Type: "has_title"})
		}
	}
	return graph
}

func buildSheetProfile(ctx context.Context, workbookID string, sheet workbook.Sheet, options Options) (SheetProfile, error) {
	rowCount := len(sheet.Rows)
	columnCount := workbook.MaxColumnCount(sheet.Rows)
	headerRowIdx := detectHeaderRow(sheet.Rows)
	profile := SheetProfile{
		Name:        sheet.Name,
		RowCount:    rowCount,
		ColumnCount: columnCount,
		HeaderRow:   headerRowIdx + 1,
		Columns:     make([]ColumnProfile, 0, columnCount),
	}
	if rowCount > 0 && columnCount > 0 {
		start, _ := excelize.CoordinatesToCellName(1, headerRowIdx+1)
		end, _ := excelize.CoordinatesToCellName(columnCount, rowCount)
		profile.TableRegions = append(profile.TableRegions, TableRegion{
			Address:  start + ":" + end,
			StartRow: headerRowIdx + 1,
			EndRow:   rowCount,
			StartCol: 1,
			EndCol:   columnCount,
		})
	}

	headers := []string{}
	if headerRowIdx >= 0 && headerRowIdx < rowCount {
		headers = append(headers, sheet.Rows[headerRowIdx]...)
	}
	for colIdx := 0; colIdx < columnCount; colIdx++ {
		name := headerName(headers, colIdx)
		column, err := buildColumnProfile(ctx, workbookID, sheet.Name, sheet.Rows, headerRowIdx, colIdx, name, options)
		if err != nil {
			return SheetProfile{}, err
		}
		profile.Columns = append(profile.Columns, column)
	}
	profile.TitleRows = titleRows(workbookID, sheet, headerRowIdx, defaultSampleRows)
	profile.SampleRows = sampleRows(workbookID, sheet, headerRowIdx, defaultSampleRows)
	return profile, nil
}

func buildColumnProfile(ctx context.Context, workbookID, sheetName string, rows [][]string, headerRowIdx, colIdx int, name string, options Options) (ColumnProfile, error) {
	letter, _ := excelize.ColumnNumberToName(colIdx + 1)
	values := make([]string, 0, len(rows))
	nonEmpty := 0
	distinct := make(map[string]int)
	var sampleValues []string
	var numeric []float64

	for rowIdx := headerRowIdx + 1; rowIdx < len(rows); rowIdx++ {
		value := strings.TrimSpace(cell(rows, rowIdx, colIdx))
		values = append(values, value)
		if value == "" {
			continue
		}
		nonEmpty++
		distinct[value]++
		if len(sampleValues) < 3 {
			sampleValues = append(sampleValues, value)
		}
		if number, ok := parseNumber(value); ok {
			numeric = append(numeric, number)
		}
	}

	inferredType := inferType(values)
	profile := ColumnProfile{
		Name:                  name,
		Index:                 colIdx + 1,
		Letter:                letter,
		InferredType:          inferredType,
		SampleValues:          sampleValues,
		NonEmptyCount:         nonEmpty,
		DistinctCountEstimate: len(distinct),
		ValueHints:            topValuePatterns(distinct, 3),
	}
	if len(values) > 0 {
		profile.EmptyRatio = float64(len(values)-nonEmpty) / float64(len(values))
	}
	if len(numeric) > 0 {
		profile.NumericStats = numericStats(numeric)
	}
	if options.ColumnTagger != nil {
		tags, err := options.ColumnTagger.Tags(ctx, ColumnTagRequest{
			WorkbookID:   workbookID,
			SheetName:    sheetName,
			ColumnName:   name,
			ColumnIndex:  colIdx + 1,
			ColumnLetter: letter,
			InferredType: inferredType,
			SampleValues: append([]string(nil), sampleValues...),
			ValueHints:   append([]ValuePattern(nil), profile.ValueHints...),
		})
		if err != nil {
			return ColumnProfile{}, err
		}
		profile.SemanticTags = uniqueStrings(tags)
	}
	return profile, nil
}

func titleRows(workbookID string, sheet workbook.Sheet, headerRowIdx, maxRows int) []EvidenceRow {
	if maxRows <= 0 || headerRowIdx <= 0 {
		return nil
	}
	out := make([]EvidenceRow, 0, maxRows)
	for rowIdx := 0; rowIdx < headerRowIdx && len(out) < maxRows; rowIdx++ {
		if rowNonEmptyCount(sheet.Rows[rowIdx]) == 0 {
			continue
		}
		out = append(out, buildEvidenceRow(workbookID, sheet, rowIdx, -1, 0, nil))
	}
	return out
}

func sampleRows(workbookID string, sheet workbook.Sheet, headerRowIdx, maxRows int) []EvidenceRow {
	if maxRows <= 0 || len(sheet.Rows) <= headerRowIdx+1 {
		return nil
	}
	out := make([]EvidenceRow, 0, maxRows)
	for rowIdx := headerRowIdx + 1; rowIdx < len(sheet.Rows) && len(out) < maxRows; rowIdx++ {
		if rowNonEmptyCount(sheet.Rows[rowIdx]) == 0 {
			continue
		}
		out = append(out, buildEvidenceRow(workbookID, sheet, rowIdx, headerRowIdx, 0, nil))
	}
	return out
}

func buildEvidenceRow(workbookID string, sheet workbook.Sheet, rowIdx, headerRowIdx int, score float64, matched []string) EvidenceRow {
	headers := []string{}
	if headerRowIdx >= 0 && headerRowIdx < len(sheet.Rows) {
		headers = append(headers, sheet.Rows[headerRowIdx]...)
	}
	cells := make(map[string]string)
	maxCols := workbook.MaxColumnCount(sheet.Rows)
	for colIdx := 0; colIdx < maxCols; colIdx++ {
		key := headerName(headers, colIdx)
		cells[key] = cell(sheet.Rows, rowIdx, colIdx)
	}
	start, _ := excelize.CoordinatesToCellName(1, rowIdx+1)
	end, _ := excelize.CoordinatesToCellName(maxCols, rowIdx+1)
	return EvidenceRow{
		Sheet:          sheet.Name,
		RowIndex:       rowIdx + 1,
		AddressRange:   start + ":" + end,
		Score:          score,
		MatchedColumns: append([]string(nil), matched...),
		Cells:          cells,
	}
}

func fingerprintWorkbook(book *workbook.Workbook) string {
	hash := sha256.New()
	for _, sheet := range book.Sheets {
		_, _ = hash.Write([]byte(sheet.Name))
		_, _ = hash.Write([]byte(fmt.Sprintf(":%d:%d", len(sheet.Rows), workbook.MaxColumnCount(sheet.Rows))))
		for rowIdx, row := range sheet.Rows {
			if rowIdx > 5 && rowIdx < len(sheet.Rows)-2 {
				continue
			}
			_, _ = hash.Write([]byte(strings.Join(row, "\x00")))
		}
	}
	return hex.EncodeToString(hash.Sum(nil))
}

func detectHeaderRow(rows [][]string) int {
	if len(rows) == 0 {
		return 0
	}
	limit := len(rows)
	if limit > 30 {
		limit = 30
	}
	bestIdx := 0
	bestScore := -1.0
	for rowIdx := 0; rowIdx < limit; rowIdx++ {
		row := rows[rowIdx]
		nonEmpty := rowNonEmptyCount(row)
		if nonEmpty == 0 {
			continue
		}
		unique := uniqueNonEmptyCount(row)
		nextNonEmpty := 0
		if rowIdx+1 < len(rows) {
			nextNonEmpty = rowNonEmptyCount(rows[rowIdx+1])
		}
		score := float64(nonEmpty*2 + unique)
		if nextNonEmpty >= nonEmpty && nonEmpty >= 2 {
			score += 4
		}
		if nonEmpty == 1 {
			score -= 4
		}
		if rowIdx == 0 && nonEmpty >= 2 {
			score += 1
		}
		if score > bestScore {
			bestScore = score
			bestIdx = rowIdx
		}
	}
	return bestIdx
}

func rowNonEmptyCount(row []string) int {
	count := 0
	for _, value := range row {
		if strings.TrimSpace(value) != "" {
			count++
		}
	}
	return count
}

func uniqueNonEmptyCount(row []string) int {
	seen := make(map[string]bool)
	for _, value := range row {
		value = strings.TrimSpace(value)
		if value != "" {
			seen[value] = true
		}
	}
	return len(seen)
}

func inferType(values []string) string {
	nonEmpty := 0
	numberCount := 0
	boolCount := 0
	dateLikeCount := 0
	for _, value := range values {
		value = strings.TrimSpace(value)
		if value == "" {
			continue
		}
		nonEmpty++
		if _, ok := parseNumber(value); ok {
			numberCount++
			continue
		}
		lower := strings.ToLower(value)
		if lower == "true" || lower == "false" || value == "是" || value == "否" {
			boolCount++
			continue
		}
		if looksDateLike(value) {
			dateLikeCount++
		}
	}
	if nonEmpty == 0 {
		return "empty"
	}
	if numberCount == nonEmpty {
		return "number"
	}
	if boolCount == nonEmpty {
		return "boolean"
	}
	if dateLikeCount == nonEmpty {
		return "date_like"
	}
	if numberCount > 0 || boolCount > 0 || dateLikeCount > 0 {
		return "mixed"
	}
	return "string"
}

func numericStats(values []float64) *NumericStats {
	stats := &NumericStats{Count: float64(len(values)), Min: values[0], Max: values[0]}
	var sum float64
	for _, value := range values {
		if value < stats.Min {
			stats.Min = value
		}
		if value > stats.Max {
			stats.Max = value
		}
		sum += value
	}
	stats.Mean = sum / float64(len(values))
	return stats
}

func topValuePatterns(counts map[string]int, limit int) []ValuePattern {
	values := make([]ValuePattern, 0, len(counts))
	for value, count := range counts {
		values = append(values, ValuePattern{Value: value, Count: count})
	}
	sort.Slice(values, func(i, j int) bool {
		if values[i].Count == values[j].Count {
			return values[i].Value < values[j].Value
		}
		return values[i].Count > values[j].Count
	})
	if len(values) > limit {
		values = values[:limit]
	}
	return values
}

func columnSummary(column ColumnProfile) string {
	parts := []string{column.InferredType}
	if len(column.SemanticTags) > 0 {
		parts = append(parts, "tags="+strings.Join(column.SemanticTags, ","))
	}
	if column.NonEmptyCount > 0 {
		parts = append(parts, fmt.Sprintf("%d non-empty", column.NonEmptyCount))
	}
	return strings.Join(parts, "; ")
}

func headerName(headers []string, colIdx int) string {
	if colIdx >= 0 && colIdx < len(headers) {
		if value := strings.TrimSpace(headers[colIdx]); value != "" {
			return value
		}
	}
	letter, _ := excelize.ColumnNumberToName(colIdx + 1)
	return letter
}

func cell(rows [][]string, rowIdx, colIdx int) string {
	if rowIdx < 0 || rowIdx >= len(rows) || colIdx < 0 || colIdx >= len(rows[rowIdx]) {
		return ""
	}
	return rows[rowIdx][colIdx]
}

func parseNumber(value string) (float64, bool) {
	text := strings.TrimSpace(value)
	text = strings.ReplaceAll(text, ",", "")
	text = strings.ReplaceAll(text, "，", "")
	text = strings.ReplaceAll(text, "￥", "")
	text = strings.ReplaceAll(text, "$", "")
	if text == "" {
		return 0, false
	}
	num, err := strconv.ParseFloat(text, 64)
	if err != nil || math.IsNaN(num) || math.IsInf(num, 0) {
		return 0, false
	}
	return num, true
}

func looksDateLike(value string) bool {
	value = strings.TrimSpace(value)
	if len(value) < 8 {
		return false
	}
	return strings.Contains(value, "-") || strings.Contains(value, "/") || strings.Contains(value, "年")
}

func nodeID(parts ...string) string {
	escaped := make([]string, len(parts))
	for i, part := range parts {
		part = strings.TrimSpace(part)
		part = strings.ReplaceAll(part, " ", "_")
		part = strings.ReplaceAll(part, "!", "_")
		part = strings.ReplaceAll(part, ".", "_")
		escaped[i] = part
	}
	return strings.Join(escaped, ":")
}

func rowIndexes(rows []EvidenceRow) []int {
	indexes := make([]int, 0, len(rows))
	for _, row := range rows {
		indexes = append(indexes, row.RowIndex)
	}
	return indexes
}
