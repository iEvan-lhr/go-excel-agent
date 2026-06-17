package engine

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// ExportMarkdown exports all sheets in the workbook as markdown files in the specified directory.
func (e *Engine) ExportMarkdown(ctx context.Context, outputDir string) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	if e.Book == nil {
		return fmt.Errorf("工作簿为空，无法导出")
	}
	if outputDir == "" {
		return fmt.Errorf("导出目录路径不能为空")
	}

	// Create directory if it does not exist
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("无法创建导出目录 '%s': %w", outputDir, err)
	}

	usedFilenames := make(map[string]bool)

	for _, sheet := range e.Book.Sheets {
		if err := ctx.Err(); err != nil {
			return err
		}

		safeBase := sanitizeFilename(sheet.Name)
		filename := safeBase + ".md"

		// Resolve collisions (case-insensitive)
		counter := 1
		for usedFilenames[strings.ToLower(filename)] {
			filename = fmt.Sprintf("%s_%d.md", safeBase, counter)
			counter++
		}
		usedFilenames[strings.ToLower(filename)] = true

		// Render sheet to markdown
		mdContent := sheet.ToMarkdown()
		outputPath := filepath.Join(outputDir, filename)

		if err := os.WriteFile(outputPath, []byte(mdContent), 0644); err != nil {
			return fmt.Errorf("写入 Markdown 文件 '%s' 失败: %w", outputPath, err)
		}
	}

	return nil
}

// sanitizeFilename replaces invalid filename characters in Windows with underscore.
func sanitizeFilename(name string) string {
	invalidChars := []string{"\\", "/", ":", "*", "?", "\"", "<", ">", "|"}
	sanitized := name
	for _, char := range invalidChars {
		sanitized = strings.ReplaceAll(sanitized, char, "_")
	}
	sanitized = strings.TrimSpace(sanitized)
	if sanitized == "" {
		sanitized = "sheet"
	}
	return sanitized
}
