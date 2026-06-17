package engine

import (
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/iEvan-lhr/go-excel-agent/workbook"
)

func TestSheetToMarkdown(t *testing.T) {
	// 1. Empty Sheet Test
	emptySheet := workbook.Sheet{
		Name: "Empty",
		Rows: [][]string{},
	}
	mdEmpty := emptySheet.ToMarkdown()
	if !strings.Contains(mdEmpty, "# Empty") || !strings.Contains(mdEmpty, "Empty sheet.") {
		t.Errorf("Expected empty sheet message, got:\n%s", mdEmpty)
	}

	// 2. Normal Sheet Test
	normalSheet := workbook.Sheet{
		Name: "Normal",
		Rows: [][]string{
			{"Name", "Price"},
			{"Apple", "1.2"},
			{"Banana", "0.8"},
		},
	}
	mdNormal := normalSheet.ToMarkdown()
	expectedLines := []string{
		"# Normal",
		"| Name | Price |",
		"| --- | --- |",
		"| Apple | 1.2 |",
		"| Banana | 0.8 |",
	}
	for _, expected := range expectedLines {
		if !strings.Contains(mdNormal, expected) {
			t.Errorf("Expected markdown to contain %q, but got:\n%s", expected, mdNormal)
		}
	}

	// 3. Special Characters Escape Test
	specialSheet := workbook.Sheet{
		Name: "Special",
		Rows: [][]string{
			{"Col|1", "Col\\2"},
			{"Line1\nLine2", "Pipe | Slash \\"},
		},
	}
	mdSpecial := specialSheet.ToMarkdown()
	expectedEscaped := []string{
		"Col\\|1",
		"Col\\\\2",
		"Line1<br>Line2",
		"Pipe \\| Slash \\\\",
	}
	for _, expected := range expectedEscaped {
		if !strings.Contains(mdSpecial, expected) {
			t.Errorf("Expected markdown to contain %q, but got:\n%s", expected, mdSpecial)
		}
	}
}

func TestEngineExportMarkdown(t *testing.T) {
	e := New()
	// Set up sheets: one is Normal, two have names that conflict under sanitization ("A/B" and "A\B")
	if err := e.LoadSheets(context.Background(), map[string][][]any{
		"A/B": {
			{"ID", "Val"},
			{"1", "X"},
		},
		"A\\B": {
			{"ID", "Val"},
			{"2", "Y"},
		},
		"EmptySheet": {},
	}); err != nil {
		t.Fatalf("Failed to load sheets: %v", err)
	}

	// Temp output directory
	tempDir, err := os.MkdirTemp("", "go-excel-agent-export-test-*")
	if err != nil {
		t.Fatalf("Failed to create temp dir: %v", err)
	}
	defer os.RemoveAll(tempDir)

	// Export to tempDir
	err = e.ExportMarkdown(context.Background(), tempDir)
	if err != nil {
		t.Fatalf("Failed to export markdown: %v", err)
	}

	// Check files created
	// A/B -> A_B.md
	// A\B -> A_B_1.md
	// EmptySheet -> EmptySheet.md
	files, err := os.ReadDir(tempDir)
	if err != nil {
		t.Fatalf("Failed to read temp dir: %v", err)
	}

	expectedFiles := map[string]bool{
		"A_B.md":        false,
		"A_B_1.md":      false,
		"EmptySheet.md": false,
	}

	for _, file := range files {
		name := file.Name()
		if _, expected := expectedFiles[name]; expected {
			expectedFiles[name] = true
		} else {
			t.Errorf("Unexpected file created: %s", name)
		}
	}

	for name, found := range expectedFiles {
		if !found {
			t.Errorf("Expected file not found: %s", name)
		}
	}

	// Validate contents of A_B.md
	contentAB, err := os.ReadFile(filepath.Join(tempDir, "A_B.md"))
	if err != nil {
		t.Fatalf("Failed to read A_B.md: %v", err)
	}
	if !strings.Contains(string(contentAB), "| 1 | X |") {
		t.Errorf("Expected content in A_B.md, got:\n%s", string(contentAB))
	}

	// Validate contents of A_B_1.md
	contentAB1, err := os.ReadFile(filepath.Join(tempDir, "A_B_1.md"))
	if err != nil {
		t.Fatalf("Failed to read A_B_1.md: %v", err)
	}
	if !strings.Contains(string(contentAB1), "| 2 | Y |") {
		t.Errorf("Expected content in A_B_1.md, got:\n%s", string(contentAB1))
	}
}
