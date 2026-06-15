# Go Excel Agent

`go-excel-agent` 是一个面向 Go 项目的 Excel 操作与 Agent 记忆内核。它提供稳定的 Excel API、强类型 DSL、操作分级、状态机日志、向量经验规范、保样式保存、Diff 追踪、工作簿结构索引、上下文胶囊和操作记忆能力。

这个仓库不绑定任何大模型供应商，也不保存用户的大模型 token。它更适合作为下游 Web 服务、CLI、自动化任务或 Agent 应用的 Excel 服务内核。外部项目可以自行接入 `go-llm-client`、OpenAI、DashScope、DeepSeek 或私有化模型。

README 中的示例均使用虚构的通用库存数据，仅用于演示 API 调用。

## 核心能力

- 打开和读取 `.xlsx` 工作簿
- 获取 workbook summary、sheet profile、行列和范围内容
- 按内容搜索单元格或行
- 修改单元格、批量修改范围或搜索结果
- 创建 sheet、清空单元格、插入空白单元格并移动原数据
- 支持数值聚合：`SUM`、`AVERAGE`、`COUNT`
- 保存时尽量保留原 Excel 样式
- 根据 Go 值类型自动选择正确的 excelize 写入方法
- 返回 `Diff`，明确知道哪些单元格或结构被修改
- 提供强类型 DSL，便于大模型输出和外部系统调用
- 提供操作注册表，记录每个操作的等级、风险和确认策略
- 提供交互状态机 `HandleInteraction`，调用方可以获取每一步状态日志
- 构建上下文胶囊 `ContextCapsule`，避免把大文件整表塞进模型上下文
- 记录操作账本、执行摘要和 session focus，方便多轮对话接续
- 提供向量经验接口，支持把成功/失败操作沉淀为可检索经验
- 预留模型接口：可注入列语义识别、请求泛化、执行总结等能力

## 安装

```bash
go get github.com/iEvan-lhr/go-excel-agent
```

## 快速开始

```go
package main

import (
	"context"
	"fmt"

	excelagent "github.com/iEvan-lhr/go-excel-agent"
)

func main() {
	ctx := context.Background()

	book, err := excelagent.Open(ctx, "demo.xlsx")
	if err != nil {
		panic(err)
	}

	summary, err := book.Summary(ctx)
	if err != nil {
		panic(err)
	}
	fmt.Printf("summary: %#v\n", summary)

	found, err := book.Find(ctx, excelagent.FindRequest{
		Sheet:        "库存台账",
		Type:         "search",
		Query:        "标准键盘",
		SearchColumn: "品名",
	})
	if err != nil {
		panic(err)
	}
	fmt.Printf("find result: %#v\n", found)

	rows, err := book.GetRange(ctx, excelagent.RangeRequest{
		Sheet: "库存台账",
		Range: "A1:D10",
	})
	if err != nil {
		panic(err)
	}
	fmt.Printf("range rows: %#v\n", rows)

	diff, err := book.UpdateCell(ctx, excelagent.UpdateCellRequest{
		Sheet: "库存台账",
		Cell:  "D2",
		Value: 128.5,
	})
	if err != nil {
		panic(err)
	}
	fmt.Printf("changed cells: %d\n", diff.ChangedCells)

	if err := book.SaveAs(ctx, "demo-output.xlsx"); err != nil {
		panic(err)
	}
}
```

## 对外 API

常用入口：

```go
book, err := excelagent.Open(ctx, "demo.xlsx")
```

读取结构：

```go
summary, err := book.Summary(ctx)
capsule, err := book.BuildContextCapsule(ctx, excelagent.ContextRequest{
	Purpose:  excelagent.PurposeUnderstandFile,
	Query:    "这个文件主要包含什么？",
	MaxRows:  20,
	MaxCells: 200,
	MaxNodes: 40,
})
```

查询数据：

```go
found, err := book.Find(ctx, excelagent.FindRequest{
	Sheet:        "库存台账",
	Type:         "search",
	Query:        "标准键盘",
	SearchColumn: "品名",
})

rows, err := book.GetRange(ctx, excelagent.RangeRequest{
	Sheet: "库存台账",
	Range: "A1:D10",
})
```

修改和保存：

```go
diff, err := book.UpdateCell(ctx, excelagent.UpdateCellRequest{
	Sheet: "库存台账",
	Cell:  "D2",
	Value: 128.5,
})

err = book.SaveAs(ctx, "out.xlsx")
```

结构和单元格操作：

```go
diff, err := book.CreateSheet(ctx, excelagent.CreateSheetRequest{
	Sheet: "销售汇总",
})
if err != nil {
	return err
}
fmt.Println(diff.StructureChanges)

diff, err = book.ClearCell(ctx, excelagent.ClearCellRequest{
	Sheet: "库存台账",
	Cell:  "D2",
})

diff, err = book.InsertCells(ctx, excelagent.InsertCellsRequest{
	Sheet: "库存台账",
	Cell:  "B2",
	Shift: "right", // right 或 down
})
```

批量更新：

```go
diff, err := book.BatchUpdate(ctx, excelagent.BatchUpdateRequest{
	Sheet: "库存台账",
	Scope: excelagent.Scope{
		Type:         "search",
		Query:        "标准键盘",
		SearchColumn: "品名",
	},
	TargetColumn: "单价",
	Action: excelagent.UpdateAction{
		Type:  "multiply",
		Value: 0.9,
	},
})
```

聚合计算：

```go
avg, err := book.Aggregate(ctx, excelagent.AggregateRequest{
	Sheet:  "库存台账",
	Column: "单价",
	Type:   "AVERAGE",
})
```

支持的聚合类型：

- `SUM`
- `AVERAGE`
- `COUNT`

## 强类型 DSL

除了直接调用 Go 方法，也可以让上层 Agent 生成统一 DSL，然后交给 `book.Execute`。

```go
_, diff, err := book.Execute(ctx, excelagent.Command{
	Op: "update_cell",
	Target: excelagent.Target{
		Sheet: "库存台账",
		Cell:  "D2",
	},
	Args: excelagent.UpdateCellArgs{Value: 128.5},
})
```

模型输出的 JSON 可以是：

```json
{
  "op": "update_cell",
  "target": {
    "sheet": "库存台账",
    "cell": "D2"
  },
  "args": {
    "value": 128.5
  }
}
```

当前支持的 `op`：

- `inspect_workbook`
- `find`
- `get_range`
- `update_cell`
- `clear_cell`
- `create_sheet`
- `insert_cells`
- `batch_update`
- `aggregate`

结构操作示例：

```json
{
  "op": "create_sheet",
  "target": {
    "sheet": "销售汇总"
  }
}
```

```json
{
  "op": "insert_cells",
  "target": {
    "sheet": "库存台账",
    "cell": "B2"
  },
  "args": {
    "shift": "right"
  }
}
```

`clear_cell` 和 `insert_cells` 的语义不同：

- `clear_cell`：清空已有单元格内容，不移动周围数据
- `insert_cells`：插入空白单元格，并让原数据向右或向下移动

当前支持的更新动作：

- `overwrite`
- `append_suffix`
- `prepend_prefix`
- `find_and_replace`
- `multiply`

## 范围语法

`range` 支持常见 Excel 坐标：

```text
A1        单个单元格
D2:D10    单列多行
A:D       多列
3:8       多行
B2:E10    矩形区域
D         整列
7         整行
```

## 记忆与上下文胶囊

`go-excel-agent` 不把完整大文件长期塞进模型上下文。它会先建立结构索引和轻量图谱，再按本次问题构造 `ContextCapsule`。

主要记忆层：

- `ArtifactMemory`：文件级结构，包含 sheet 数量、行列规模、表头行、表格区域
- `SheetProfile`：sheet 画像，包含表头、列画像、标题行、代表行
- `DataGraph`：轻量节点网络，包含 workbook、sheet、table region、column、row cluster、operation 节点
- `OperationRecord`：精确操作账本，保存执行过的 DSL JSON、位置、Diff、错误
- `ExecutionSummary`：按需生成的执行摘要，简单修改只记录位置、原值、新值
- `SessionFocus`：当前会话焦点，保存最近文件、最近操作、最近相关行
- `ContextCapsule`：本轮给模型看的最小上下文包

构造上下文：

```go
capsule, err := book.BuildContextCapsule(ctx, excelagent.ContextRequest{
	Purpose:  excelagent.PurposePlanUpdate,
	Query:    "把标准键盘的单价打九折",
	MaxRows:  20,
	MaxCells: 200,
	MaxNodes: 40,
})
```

常用 purpose：

- `PurposeUnderstandFile`：理解文件内容
- `PurposeLocateTarget`：定位目标行、列、范围
- `PurposePlanUpdate`：准备生成修改类 DSL
- `PurposeValidate`：校验命令
- `PurposeRepair`：根据错误修复命令
- `PurposeExplainResult`：解释执行结果
- `PurposeFollowup`：处理多轮追问

执行并记录记忆：

```go
_, diff, record, err := book.ExecuteAndRemember(ctx, "把标准键盘的单价改为 128.5", excelagent.Command{
	Op: "update_cell",
	Target: excelagent.Target{
		Sheet: "库存台账",
		Cell:  "D2",
	},
	Args: excelagent.UpdateCellArgs{Value: 128.5},
})

fmt.Println(diff.ChangedCells)
fmt.Println(record.OperationID)
```

## 交互状态机与向量经验

除了直接执行 DSL，也可以使用 `HandleInteraction`。它会把一次用户请求拆成可观测的状态机流程，并在执行成功后把操作沉淀为向量经验。

典型流程：

```text
received
  -> context_built
  -> vector_retrieved
  -> operation_matched
  -> args_extracted
  -> validated
  -> need_confirmation / executable
  -> executing
  -> executed
  -> remembered
```

示例：先规划并等待确认。

```go
vector := excelagent.NewInMemoryVectorStore()

result, err := book.HandleInteraction(ctx, excelagent.InteractionRequest{
	WorkbookID:  "current",
	UserRequest: "创建一个销售汇总表",
}, excelagent.InteractionOptions{
	VectorStore: vector,
	EventSink: excelagent.StateEventSinkFunc(func(ctx context.Context, event excelagent.StateEvent) error {
		fmt.Printf("%s: %s\n", event.State, event.Message)
		return nil
	}),
})
if err != nil {
	return err
}

if result.Status == excelagent.StateNeedConfirmation {
	fmt.Println(result.Message)
}
```

确认后执行同一个 command：

```go
result, err = book.HandleInteraction(ctx, excelagent.InteractionRequest{
	WorkbookID:  "current",
	UserRequest: "创建一个销售汇总表",
	Command:     result.Command,
	Confirmed:   true,
}, excelagent.InteractionOptions{
	VectorStore: vector,
})
if err != nil {
	return err
}

fmt.Println(result.Status)
fmt.Println(result.Diff.StructureChanges)
```

`InteractionResult.Events` 会返回完整状态日志。调用方可以把它用于前端进度、调试、审计或后台日志。

### 操作分级

内置操作注册表记录每个操作的等级、风险和确认策略：

```go
registry := excelagent.BuiltinOperationRegistry()
spec, ok := registry.Get("create_sheet")
fmt.Println(ok, spec.Level, spec.Risk, spec.ConfirmationPolicy)
```

常见等级：

- `LevelReadOnly`：只读操作，例如 `get_range`、`aggregate`
- `LevelLocate`：定位操作，例如 `find`
- `LevelCellEdit`：单元格修改，例如 `update_cell`、`clear_cell`
- `LevelRangeEdit`：批量范围修改，例如 `batch_update`
- `LevelStructureEdit`：结构修改，例如 `create_sheet`、`insert_cells`
- `LevelDestructive`：删除类高风险操作
- `LevelExternalWrite`：保存、导出等外部写入

默认策略下，结构修改、删除类操作和外部写入会进入确认状态。调用方可以通过 `InteractionOptions.RequireConfirmationForLevels` 调整策略。

### 向量经验

`VectorStore` 是一个接口，不绑定具体向量数据库。当前提供 `NewInMemoryVectorStore`，便于本地测试和服务内缓存。

```go
type VectorStore interface {
	Upsert(ctx context.Context, records []excelagent.VectorRecord) error
	Search(ctx context.Context, req excelagent.VectorSearchRequest) ([]excelagent.VectorSearchResult, error)
}
```

一次操作成功后，系统会把 `OperationRecord` 转成 `VectorRecord`：

```text
用户原话 -> 实际 op -> 参数模式 -> 成功/失败结果 -> 可检索经验
```

下一次相似请求进来时，`HandleInteraction` 会先检索操作规范和历史经验。如果相似经验的 `op` 一致、参数完整、风险可接受，就可以快速生成 command；如果候选操作冲突或参数不足，则进入 `need_clarification`。

## 模型接口与自定义扩展

这个仓库只定义模型抽象，不绑定具体模型客户端。

核心接口：

```go
type TextModel interface {
	Complete(ctx context.Context, req excelagent.ModelRequest) (excelagent.ModelResponse, error)
}
```

可注入扩展点：

```go
book, err := excelagent.OpenWithMemoryOptions(ctx, "demo.xlsx",
	excelagent.WithColumnTagger(myColumnTagger),
	excelagent.WithIntentGeneralizer(myIntentGeneralizer),
	excelagent.WithExecutionSummarizer(mySummarizer),
)
```

扩展点说明：

- `ColumnTagger`：列语义标签识别。默认不写死任何业务词，也不自动猜测业务语义
- `IntentGeneralizer`：把用户请求和执行命令泛化成结构化任务类型
- `ExecutionSummarizer`：把执行结果总结成长期记忆

如果你想用规则，也可以在外部传入规则：

```go
tagger := excelagent.RuleBasedColumnTagger{
	Rules: []excelagent.SemanticTagRule{
		{Tag: "item_name", Terms: []string{"品名", "名称"}},
		{Tag: "unit_price", Terms: []string{"单价", "价格"}},
	},
}

book, err := excelagent.OpenWithMemoryOptions(ctx, "demo.xlsx",
	excelagent.WithColumnTagger(tagger),
)
```

如果你想用大模型，也可以实现 `ColumnTagger`，或使用 `ModelColumnTagger`：

```go
tagger := excelagent.ModelColumnTagger{
	Model: myModel,
}

book, err := excelagent.OpenWithMemoryOptions(ctx, "demo.xlsx",
	excelagent.WithColumnTagger(tagger),
)
```

## 在另一个 Web 仓库中接入

推荐把另一个仓库设计成 Web/Agent 应用层，把本仓库作为 Excel 内核。

典型职责划分：

```text
你的 Web 仓库
  - 用户登录和会话
  - 文件上传和文件存储
  - 用户模型 token / provider / model 配置
  - 调用大模型
  - 管理前端对话流
  - 调用 go-excel-agent 执行 Excel 操作

go-excel-agent
  - 打开 Excel
  - 构建 workbook memory
  - 构建 context capsule
  - 检索操作规范和向量经验
  - 输出状态机日志
  - 执行 DSL
  - 返回 diff
  - 保样式保存
  - 记录操作账本、向量经验和 session focus
```

用户上传文件后的服务端流程：

```go
func HandleUpload(ctx context.Context, path string) (*excelagent.Book, error) {
	book, err := excelagent.Open(ctx, path)
	if err != nil {
		return nil, err
	}
	return book, nil
}
```

用户发起一次对话：

```go
func HandleChat(ctx context.Context, book *excelagent.Book, userText string, model excelagent.TextModel) (string, error) {
	capsule, err := book.BuildContextCapsule(ctx, excelagent.ContextRequest{
		Purpose:  excelagent.PurposePlanUpdate,
		Query:    userText,
		MaxRows:  20,
		MaxCells: 200,
		MaxNodes: 40,
	})
	if err != nil {
		return "", err
	}

	prompt := RenderPrompt(userText, capsule)
	resp, err := model.Complete(ctx, excelagent.ModelRequest{
		SystemPrompt: "你是 Excel Agent。只基于给定上下文生成合法 DSL 或回答。",
		Prompt:       prompt,
	})
	if err != nil {
		return "", err
	}

	return resp.Content, nil
}
```

如果模型输出 DSL，你的 Web 仓库负责 JSON 解析和校验，然后执行：

```go
var cmd excelagent.Command
if err := json.Unmarshal([]byte(modelOutput), &cmd); err != nil {
	return err
}

_, diff, record, err := book.ExecuteAndRemember(ctx, userText, cmd)
if err != nil {
	return err
}

fmt.Println(diff.ChangedCells)
fmt.Println(record.OperationID)
```

如果你希望优先走操作经验检索和状态机日志，可以让 Web 仓库调用 `HandleInteraction`：

```go
vector := excelagent.NewInMemoryVectorStore()

result, err := book.HandleInteraction(ctx, excelagent.InteractionRequest{
	WorkbookID:  "current",
	UserRequest: userText,
}, excelagent.InteractionOptions{
	VectorStore: vector,
	EventSink: excelagent.StateEventSinkFunc(func(ctx context.Context, event excelagent.StateEvent) error {
		// 可以推送给前端，也可以写入服务端日志。
		return nil
	}),
})
if err != nil {
	return err
}

switch result.Status {
case excelagent.StateNeedClarification:
	return fmt.Errorf("%s", result.Message)
case excelagent.StateNeedConfirmation:
	// 把 result.Command 暂存到会话，等待用户确认后再次调用 HandleInteraction。
	return nil
case excelagent.StateRemembered:
	fmt.Println(result.Diff)
}
```

多轮对话时，不建议把历史 3000 行继续塞回 prompt。推荐保存：

```text
user session
  user_id
  workbook_id
  uploaded_file_path
  selected provider/model
  encrypted or external-managed token reference
  chat messages
  book.Memory().Session
  operation ids
```

下一轮请求只需要重新调用：

```go
capsule, err := book.BuildContextCapsule(ctx, excelagent.ContextRequest{
	Purpose:           excelagent.PurposeFollowup,
	Query:             userText,
	IncludeOperations: true,
	MaxRows:           20,
	MaxCells:          200,
	MaxNodes:          40,
})
```

## 对接 go-llm-client 示例

另一个仓库可以把 `go-llm-client` 包装成 `excelagent.TextModel`。

```go
package yourapp

import (
	"context"
	"strings"

	excelagent "github.com/iEvan-lhr/go-excel-agent"
	"github.com/iEvan-lhr/go-llm-client/client"
	"github.com/iEvan-lhr/go-llm-client/llm"
)

type LLMAdapter struct {
	c *client.Client
}

func NewLLMAdapter(provider, model, apiKey, apiURL string) (*LLMAdapter, error) {
	c, err := client.New(llm.Config{
		Provider: provider,
		Model:    model,
		APIKey:   apiKey,
		APIURL:   apiURL,
	})
	if err != nil {
		return nil, err
	}
	return &LLMAdapter{c: c}, nil
}

func (a *LLMAdapter) Complete(ctx context.Context, req excelagent.ModelRequest) (excelagent.ModelResponse, error) {
	resp, err := a.c.SendNoHistory(ctx, renderPrompt(req))
	if err != nil {
		return excelagent.ModelResponse{}, err
	}
	return excelagent.ModelResponse{Content: resp.Message.Content}, nil
}

func renderPrompt(req excelagent.ModelRequest) string {
	var parts []string
	if req.SystemPrompt != "" {
		parts = append(parts, "System:\n"+req.SystemPrompt)
	}
	for _, msg := range req.Messages {
		parts = append(parts, msg.Role+":\n"+msg.Content)
	}
	if req.Prompt != "" {
		parts = append(parts, req.Prompt)
	}
	return strings.Join(parts, "\n\n")
}
```

然后在 Web 仓库里使用用户自己的配置：

```go
model, err := NewLLMAdapter(userProvider, userModel, userAPIKey, userAPIURL)
if err != nil {
	return err
}

answer, err := HandleChat(ctx, book, userText, model)
```

这个设计里，用户 token 和 provider 配置都留在你的 Web 仓库，不进入 `go-excel-agent`。

## 真实大模型集成测试

仓库提供了一个手动集成测试文件：[llm_integration_test.go](./llm_integration_test.go)。

它通过 build tag 隔离，不会影响默认测试。

运行方式：

```powershell
go get github.com/iEvan-lhr/go-llm-client

$env:EXCELAGENT_TEST_XLSX="C:\path\to\real.xlsx"
$env:EXCELAGENT_LLM_PROVIDER="dashscope"
$env:EXCELAGENT_LLM_MODEL="qwen-plus"
$env:EXCELAGENT_LLM_API_KEY="..."

go test -tags llm_integration -run TestLLMReadsRealWorkbook -v
```

可选参数：

```text
EXCELAGENT_LLM_API_URL
EXCELAGENT_TEST_QUESTION
EXCELAGENT_CONTEXT_MAX_ROWS
EXCELAGENT_CONTEXT_MAX_CELLS
EXCELAGENT_CONTEXT_MAX_NODES
```

## 包结构

```text
./
  对外 API：Open、Book、Summary、Find、GetRange、UpdateCell、CreateSheet、ClearCell、InsertCells、BatchUpdate、Aggregate、SaveAs、Execute、HandleInteraction

engine/
  DSL 类型和确定性执行器，负责 find/update/create_sheet/clear_cell/insert_cells/batch_update/aggregate

memory/
  分层记忆、结构索引、上下文胶囊、操作账本、向量经验、模型扩展接口

ops/
  操作注册表，定义操作等级、风险、确认策略、参数规范和常见用户表达

workbook/
  Excel 内存模型、Sheet、Diff、StructureChange、FindResult、typed value 缓存

excelizeadapter/
  基于 excelize 的文件读写、保样式保存、自动单元格写入
```

推荐依赖方向：

```text
外部 Web / CLI / Agent
        ↓
   excelagent API
        ↓
  ops + memory + engine
        ↓
    workbook
        ↑
 excelizeadapter
```

## 设计原则

### 不绑定大模型

本仓库只提供接口，不绑定 DashScope、OpenAI、DeepSeek 或任何私有模型。模型调用、token 管理、provider 配置应该放在外部应用层。

### Engine 确定性执行

`engine` 只接受 Go request 或 DSL command。大模型只负责生成候选 DSL，不参与文件写入细节。

### 修改必须返回 Diff

任何更新操作都应该返回 `Diff`。调用方可以根据 `ChangedCells` 和 `StructureChanges` 判断是否需要保存，也可以把 Diff 写入操作账本。

### 操作先分级，再执行

每个操作都有 `OperationSpec`，包含等级、风险、确认策略和参数说明。调用方可以根据 `Level` 和 `Risk` 决定是否自动执行、要求确认或拒绝执行。

### 状态机必须可观测

`HandleInteraction` 会为一次请求输出 `StateEvent`。调用方可以实时订阅，也可以在执行后读取 `InteractionResult.Events`，用于前端进度、调试和审计。

### 向量经验不是聊天历史

向量经验保存的是“类似请求如何变成操作”的可复用路径，而不是完整对话。它记录用户原话、实际 `op`、参数模式、成功或失败结果，用于下一次快速匹配和减少模型推理。

### 上下文不是记忆本身

长期保存的是结构索引、操作账本、摘要和引用。每轮给模型的是按需构造的 `ContextCapsule`，不是完整历史上下文。

### 保存时保留样式

当工作簿来自已有文件时，保存逻辑会重新打开原文件，只写入变化单元格，从而尽量保留原有样式、列宽、公式和 workbook 结构。

### typed value 写入

底层会根据 Go 值类型选择 excelize 的写入方法：

- `int / int64` -> `SetCellInt`
- `uint / uint64` -> `SetCellUint`
- `float32 / float64` -> `SetCellFloat`
- `bool` -> `SetCellBool`
- `string` -> `SetCellStr`
- `time.Time / time.Duration / nil` -> `SetCellValue`

这样数字不会被错误保存成字符串。

## 测试

默认测试：

```bash
go test ./...
```

真实大模型测试需要手动开启 build tag：

```bash
go test -tags llm_integration -run TestLLMReadsRealWorkbook -v
```

建议重点保留以下测试场景：

- 打开 Excel 后读取 sheet 内容
- 自动探测非首行表头
- 构造上下文胶囊时不携带整表大上下文
- 修改单元格后保存
- 保存时样式不变
- 数字写入后仍是数字
- `create_sheet` 返回结构 Diff
- `clear_cell` 只清空内容，不移动周围数据
- `insert_cells` 能按 `right` 或 `down` 插入空白单元格
- `batch_update` 能处理 `D2`、`D2:D10`、搜索结果
- 没有定位到任何单元格时返回错误
- `ExecuteAndRemember` 能记录 DSL、Diff 和精确位置
- `HandleInteraction` 能输出状态机事件、处理确认状态，并写入向量经验

## License

MIT
