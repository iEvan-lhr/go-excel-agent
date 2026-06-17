package excelagent

import (
	"context"
	"testing"
)

func TestEx(t *testing.T) {
	// 1. Test using direct API Book.ExportMarkdown
	book, err := Open(context.Background(), "testing.xlsx")
	if err != nil {
		t.Fatalf("open book failed: %v", err)
	}

	err = book.ExportMarkdown(context.Background(), "md_out_1.md")
	if err != nil {
		t.Fatalf("ExportMarkdown failed: %v", err)
	}
}
