package memory

import (
	"bytes"
	"fmt"
	"text/template"
)

type PromptLevel string

const (
	LevelIntent PromptLevel = "intent"
	LevelArgs   PromptLevel = "args"
	LevelRepair PromptLevel = "repair"
)

type PromptTemplate struct {
	SystemPrompt string
	UserTemplate string
	Description  string
}

// HierarchicalPrompts is the default set of templates.
var HierarchicalPrompts = map[PromptLevel]PromptTemplate{
	LevelIntent: {
		SystemPrompt: `你是一个 Excel Agent 意图识别与任务规划路由器。
你需要根据用户请求和当前执行历史，判断下一步要执行的 Excel 操作。

当前执行历史（已完成的步骤）：
{{.ExecutionHistory}}

支持的 Excel 算子：
- inspect_workbook: 查看工作簿 sheet 及结构信息。
- find: 根据内容查找/搜索单元格或行。
- get_range: 读取特定范围的单元格内容。
- update_cell: 更新单个单元格的值。
- clear_cell: 清空单元格内容（不移位）。
- update_style: 设置单元格样式（字体、字号、颜色、背景、对齐等）。
- write_formula: 写入计算公式（如 SUM, AVERAGE 等）。
- batch_update: 批量更新单元格数值或文本。
- create_sheet: 创建一个新的工作表。
- delete_sheet: 删除一个工作表。
- insert_cells: 插入空白单元格并移位。
- save_as: 将工作簿另存为文件。

请分析用户请求：如果是复杂复合任务，请将其拆分为多个单步操作，并结合“已完成的步骤”来输出目前的“下一步操作”。
请只输出对应的 JSON 格式，不要有任何 Markdown 块或额外解释。
示例输出：{"next_op": "create_sheet", "plan_summary": "读取库存已完成，现在创建新 Sheet", "is_finished": false}`,
		UserTemplate: `用户请求: {{.Query}}`,
	},
	LevelArgs: {
		SystemPrompt: `你是一个 Excel 强类型 DSL 生成器。
当前需要执行的操作是: {{.Op}}。
该操作的详细参数定义 (Schema) 为:
{{.OpSchema}}

当前文件的上下文胶囊（包含与查询相关的行数据，或历史步骤生成的信息）：
{{.ContextCapsule}}

请严格基于上述定义和上下文，生成符合 JSON 格式的 Command。不要包含任何 Markdown 格式包裹（如 ` + "`" + `` + "`" + `` + "`" + `json），只返回纯 JSON 字符串：
{
  "op": "{{.Op}}",
  "target": { ... },
  "args": { ... }
}`,
		UserTemplate: `用户请求: {{.Query}}`,
	},
	LevelRepair: {
		SystemPrompt: `你是一个 Excel DSL 修复专家。
上一次执行的 DSL 为:
{{.FailedCommand}}

执行报错信息为:
{{.ErrorMessage}}

请根据错误原因与上下文，输出修复后的合法 JSON 格式 DSL。不要包含任何 Markdown 包裹，仅返回纯 JSON 字符串。`,
		UserTemplate: `请帮我修复上述 DSL`,
	},
}

type IntentPromptData struct {
	ExecutionHistory string
	Query            string
}

type ArgsPromptData struct {
	Op             string
	OpSchema       string
	ContextCapsule string
	Query          string
}

type RepairPromptData struct {
	FailedCommand string
	ErrorMessage  string
}

func RenderPromptTemplate(tmpl string, data any) (string, error) {
	t, err := template.New("prompt").Parse(tmpl)
	if err != nil {
		return "", err
	}
	var buf bytes.Buffer
	if err := t.Execute(&buf, data); err != nil {
		return "", err
	}
	return buf.String(), nil
}

func (s *Store) RenderPrompt(level PromptLevel, data any) (ModelRequest, error) {
	tmpl, ok := HierarchicalPrompts[level]
	if !ok {
		return ModelRequest{}, fmt.Errorf("unknown prompt level: %s", level)
	}

	sysPrompt, err := RenderPromptTemplate(tmpl.SystemPrompt, data)
	if err != nil {
		return ModelRequest{}, err
	}

	userPrompt, err := RenderPromptTemplate(tmpl.UserTemplate, data)
	if err != nil {
		return ModelRequest{}, err
	}

	return ModelRequest{
		SystemPrompt: sysPrompt,
		Prompt:       userPrompt,
		Metadata: map[string]any{
			"level": string(level),
		},
	}, nil
}
