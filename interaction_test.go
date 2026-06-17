package excelagent

import (
	"context"
	"testing"
)

func TestHandleInteractionCreateSheetRequiresConfirmationThenRecordsExperience(t *testing.T) {
	ctx := context.Background()
	book := New()
	if err := book.engine.LoadSheets(ctx, map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", 1200},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}
	if _, err := book.IndexMemory(ctx, "wb1"); err != nil {
		t.Fatalf("index failed: %v", err)
	}

	vector := NewInMemoryVectorStore()
	var streamed []StateEvent
	result, err := book.HandleInteraction(ctx, InteractionRequest{
		WorkbookID:  "wb1",
		UserRequest: "创建一个利润分析表",
	}, InteractionOptions{
		VectorStore: vector,
		EventSink: StateEventSinkFunc(func(ctx context.Context, event StateEvent) error {
			streamed = append(streamed, event)
			return nil
		}),
	})
	if err != nil {
		t.Fatalf("handle failed: %v", err)
	}
	if result.Status != StateNeedConfirmation {
		t.Fatalf("expected confirmation, got %s with message %q", result.Status, result.Message)
	}
	if result.Command == nil || result.Command.Op != "create_sheet" || result.Command.Target.Sheet != "利润分析" {
		t.Fatalf("unexpected planned command: %#v", result.Command)
	}
	if len(result.Events) == 0 || len(streamed) != len(result.Events) {
		t.Fatalf("expected streamed events, got result=%d streamed=%d", len(result.Events), len(streamed))
	}

	result, err = book.HandleInteraction(ctx, InteractionRequest{
		WorkbookID:  "wb1",
		UserRequest: "创建一个利润分析表",
		Command:     result.Command,
		Confirmed:   true,
	}, InteractionOptions{VectorStore: vector})
	if err != nil {
		t.Fatalf("confirmed handle failed: %v", err)
	}
	if result.Status != StateRemembered {
		t.Fatalf("expected remembered, got %s", result.Status)
	}
	if book.engine.Book.SheetByName("利润分析") == nil {
		t.Fatal("expected created sheet")
	}
	if result.Diff == nil || len(result.Diff.StructureChanges) != 1 {
		t.Fatalf("expected structure diff, got %#v", result.Diff)
	}

	found, err := vector.Search(ctx, VectorSearchRequest{
		Query: "新建利润页",
		Kinds: []VectorRecordKind{VectorKindSuccessExperience},
		Ops:   []string{"create_sheet"},
		Limit: 3,
	})
	if err != nil {
		t.Fatalf("search failed: %v", err)
	}
	if len(found) == 0 {
		t.Fatal("expected stored vector experience")
	}
}

func TestHandleInteractionExecutesExplicitReadOnlyCommand(t *testing.T) {
	ctx := context.Background()
	book := New()
	if err := book.engine.LoadSheets(ctx, map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", 1200},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}
	if _, err := book.IndexMemory(ctx, "wb1"); err != nil {
		t.Fatalf("index failed: %v", err)
	}

	result, err := book.HandleInteraction(ctx, InteractionRequest{
		WorkbookID:  "wb1",
		UserRequest: "标准键盘在哪",
		Command: &Command{
			Op:     "find",
			Target: Target{Sheet: "Data", SearchQuery: "标准键盘"},
			Args:   FindArgs{Type: "search"},
		},
	}, InteractionOptions{})
	if err != nil {
		t.Fatalf("handle failed: %v", err)
	}
	if result.Status != StateRemembered {
		t.Fatalf("expected remembered, got %s", result.Status)
	}
	if len(result.Events) == 0 {
		t.Fatal("expected state events")
	}
}

func TestHandleInteractionFinish(t *testing.T) {
	ctx := context.Background()
	book := New()
	if err := book.engine.LoadSheets(ctx, map[string][][]any{
		"Data": {
			{"品名", "单价"},
			{"标准键盘", 1200},
		},
	}); err != nil {
		t.Fatalf("load failed: %v", err)
	}
	if _, err := book.IndexMemory(ctx, "wb1"); err != nil {
		t.Fatalf("index failed: %v", err)
	}

	result, err := book.HandleInteraction(ctx, InteractionRequest{
		WorkbookID:  "wb1",
		UserRequest: "完成了，结束吧",
		Command: &Command{
			Op:     "finish",
			Target: Target{Sheet: "Data"},
		},
	}, InteractionOptions{})
	if err != nil {
		t.Fatalf("handle failed: %v", err)
	}
	if result.Status != StateFinished {
		t.Fatalf("expected finished state, got %s", result.Status)
	}
}
