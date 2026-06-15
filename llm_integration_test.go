//go:build llm_integration
// +build llm_integration

package excelagent

// Run manually with a real workbook and a real LLM key:
//
//   go get github.com/iEvan-lhr/go-llm-client
//   $env:EXCELAGENT_TEST_XLSX="C:\path\to\real.xlsx"
//   $env:EXCELAGENT_LLM_PROVIDER="dashscope"
//   $env:EXCELAGENT_LLM_MODEL="qwen-plus"
//   $env:EXCELAGENT_LLM_API_KEY="..."
//   go test -tags llm_integration -run TestLLMReadsRealWorkbook -v
//
// Optional:
//   EXCELAGENT_LLM_API_URL
//   EXCELAGENT_TEST_QUESTION
//   EXCELAGENT_CONTEXT_MAX_ROWS
//   EXCELAGENT_CONTEXT_MAX_CELLS
//   EXCELAGENT_CONTEXT_MAX_NODES

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"strings"
	"testing"
	"time"

	"github.com/iEvan-lhr/go-llm-client/client"
	"github.com/iEvan-lhr/go-llm-client/llm"
)

type goLLMClientAdapter struct {
	client *client.Client
}

func (a goLLMClientAdapter) Complete(ctx context.Context, req ModelRequest) (ModelResponse, error) {
	prompt := renderModelPrompt(req)
	resp, err := a.client.SendNoHistory(ctx, prompt)
	if err != nil {
		return ModelResponse{}, err
	}
	return ModelResponse{Content: resp.Message.Content}, nil
}

func TestLLMReadsRealWorkbook(t *testing.T) {
	xlsxPath := os.Getenv("EXCELAGENT_TEST_XLSX")
	if strings.TrimSpace(xlsxPath) == "" {
		t.Skip("set EXCELAGENT_TEST_XLSX to a real .xlsx file path")
	}

	apiKey := firstNonEmptyEnv("EXCELAGENT_LLM_API_KEY", "DASHSCOPE_API_KEY", "OPENAI_API_KEY", "DEEPSEEK_API_KEY")
	if strings.TrimSpace(apiKey) == "" {
		t.Skip("set EXCELAGENT_LLM_API_KEY or a provider-specific API key env var")
	}

	provider := envOrDefault("EXCELAGENT_LLM_PROVIDER", "dashscope")
	model := envOrDefault("EXCELAGENT_LLM_MODEL", "qwen-plus")
	apiURL := os.Getenv("EXCELAGENT_LLM_API_URL")

	ctx, cancel := context.WithTimeout(context.Background(), 90*time.Second)
	defer cancel()

	c, err := client.New(llm.Config{
		Provider: provider,
		Model:    model,
		APIKey:   apiKey,
		APIURL:   apiURL,
		SystemPrompt: strings.Join([]string{
			"你是一个 Excel 数据分析助手。",
			"你只能基于用户提供的 workbook context capsule 作答。",
			"如果上下文不足，请明确说明需要更多证据，而不是编造。",
		}, "\n"),
	})
	if err != nil {
		t.Fatalf("create llm client failed: %v", err)
	}

	modelAdapter := goLLMClientAdapter{client: c}
	book, err := Open(ctx, xlsxPath)
	if err != nil {
		t.Fatalf("open workbook failed: %v", err)
	}

	capsule, err := book.BuildContextCapsule(ctx, ContextRequest{
		Purpose:  PurposeUnderstandFile,
		Query:    envOrDefault("EXCELAGENT_TEST_QUESTION", "这个 Excel 文件主要包含什么内容？有哪些主要 sheet 和关键字段？"),
		MaxRows:  envIntOrDefault("EXCELAGENT_CONTEXT_MAX_ROWS", 12),
		MaxCells: envIntOrDefault("EXCELAGENT_CONTEXT_MAX_CELLS", 120),
		MaxNodes: envIntOrDefault("EXCELAGENT_CONTEXT_MAX_NODES", 30),
	})
	if err != nil {
		t.Fatalf("build context capsule failed: %v", err)
	}

	capsuleJSON, err := json.MarshalIndent(capsule, "", "  ")
	if err != nil {
		t.Fatalf("marshal context capsule failed: %v", err)
	}

	resp, err := modelAdapter.Complete(ctx, ModelRequest{
		SystemPrompt: "请阅读结构化 Excel 上下文，并给出简洁、可追溯的中文分析。",
		Messages: []ModelMessage{
			{
				Role: "user",
				Content: strings.Join([]string{
					"下面是从真实 Excel 文件构造出来的上下文胶囊。",
					"请说明这个文件大概是什么、包含哪些主要表/字段、你判断的依据是什么。",
					"不要输出 JSON，不要假设上下文之外的数据。",
					"",
					string(capsuleJSON),
				}, "\n"),
			},
		},
		Metadata: map[string]any{
			"test":      "real_workbook_llm_read",
			"xlsx_path": xlsxPath,
			"provider":  provider,
			"model":     model,
		},
	})
	if err != nil {
		t.Fatalf("llm complete failed: %v", err)
	}
	if strings.TrimSpace(resp.Content) == "" {
		t.Fatal("llm returned empty content")
	}

	t.Logf("context budget: rows=%d/%d cells=%d/%d nodes=%d/%d",
		capsule.Budget.IncludedRows,
		capsule.Budget.MaxRows,
		capsule.Budget.IncludedCells,
		capsule.Budget.MaxCells,
		capsule.Budget.IncludedNodes,
		capsule.Budget.MaxNodes,
	)
	t.Logf("llm response:\n%s", resp.Content)
}

func renderModelPrompt(req ModelRequest) string {
	var parts []string
	if strings.TrimSpace(req.SystemPrompt) != "" {
		parts = append(parts, "System:\n"+req.SystemPrompt)
	}
	for _, msg := range req.Messages {
		parts = append(parts, strings.Title(msg.Role)+":\n"+msg.Content)
	}
	if strings.TrimSpace(req.Prompt) != "" {
		parts = append(parts, req.Prompt)
	}
	return strings.Join(parts, "\n\n")
}

func firstNonEmptyEnv(keys ...string) string {
	for _, key := range keys {
		if value := strings.TrimSpace(os.Getenv(key)); value != "" {
			return value
		}
	}
	return ""
}

func envOrDefault(key, fallback string) string {
	if value := strings.TrimSpace(os.Getenv(key)); value != "" {
		return value
	}
	return fallback
}

func envIntOrDefault(key string, fallback int) int {
	value := strings.TrimSpace(os.Getenv(key))
	if value == "" {
		return fallback
	}
	var parsed int
	if _, err := fmt.Sscanf(value, "%d", &parsed); err != nil || parsed <= 0 {
		return fallback
	}
	return parsed
}
