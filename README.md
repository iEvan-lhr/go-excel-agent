# Go Excel Agent

`go-excel-agent` 是一个面向 Go 项目的 Excel DSL 执行引擎。它把 Excel 文件读取、内容查询、范围定位、单元格修改、批量更新、聚合计算和保样式保存抽象成稳定的 Go API 与 JSON DSL，方便被 CLI、Web 服务、自动化脚本或大模型 Agent 复用。

项目目标不是把大模型绑死在 Excel 逻辑里，而是先提供一个可靠的 Excel 操作内核。大模型、命令行、Web API 都可以作为上层适配器接入。

README 中的示例均使用虚构的通用库存数据，仅用于演示 API 调用。

## 特性

- 打开和读取 `.xlsx` 工作簿
- 获取 sheet、行列和范围内容
- 按内容搜索单元格或行
- 修改单元格、批量修改范围或搜索结果
- 支持数值聚合：`SUM`、`AVERAGE`、`COUNT`
- 保存时尽量保留原 Excel 样式
- 根据 Go 值类型自动选择正确的 excelize 写入方法
- 返回 `Diff`，明确知道哪些单元格被修改
- 提供强类型 DSL，便于大模型输出和外部系统调用

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

	result, err := book.Find(ctx, excelagent.FindRequest{
		Sheet:        "库存台账",
		Type:         "search",
		Query:        "标准键盘",
		SearchColumn: "品名",
	})
	if err != nil {
		panic(err)
	}
	fmt.Printf("find result: %#v\n", result)

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

## 使用 DSL 执行

除了直接调用 Go 方法，也可以使用统一的 `Command` DSL：

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

DSL 对大模型很友好。上层 Agent 可以让模型输出如下 JSON，然后交给 `book.Execute` 执行：

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

## 批量更新

按搜索结果定位行，再修改指定列：

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

等价 DSL：

```json
{
  "op": "batch_update",
  "target": {
    "sheet": "库存台账",
    "column": "单价",
    "scope": {
      "type": "search",
      "query": "标准键盘",
      "search_column": "品名"
    }
  },
  "args": {
    "action": "multiply",
    "value": 0.9
  }
}
```

当前支持的更新动作：

- `overwrite`：覆盖为新值
- `append_suffix`：追加后缀
- `prepend_prefix`：添加前缀
- `find_and_replace`：查找替换
- `multiply`：将当前数值乘以指定倍数

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

## 聚合计算

```go
avg, err := book.Aggregate(ctx, excelagent.AggregateRequest{
	Sheet:  "库存台账",
	Column: "单价",
	Type:   "AVERAGE",
})
```

支持：

- `SUM`
- `AVERAGE`
- `COUNT`

## 包结构

```text
./
  对外 API：Open、Book、Summary、Find、GetRange、UpdateCell、BatchUpdate、Aggregate、SaveAs、Execute

workbook/
  Excel 内存模型、Sheet、Diff、FindResult、typed value 缓存

engine/
  DSL 类型和执行器，负责 find/update/batch_update/aggregate

excelizeadapter/
  基于 excelize 的文件读写、保样式保存、自动单元格写入
```

推荐依赖方向：

```text
外部应用 / Agent
        ↓
   excelagent API
        ↓
     engine
        ↓
    workbook
        ↑
 excelizeadapter
```

## 设计理念

### Engine 不依赖大模型

`engine` 是确定性执行层，只接受 Go request 或 DSL command。大模型只负责生成 DSL，不参与文件写入细节。

### 修改必须返回 Diff

任何更新操作都应该返回 `workbook.Diff`。调用方可以根据 `ChangedCells` 判断是否需要保存，避免“模型说改了但实际没改”的问题。

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

```bash
go test ./...
```

建议重点保留以下测试场景：

- 打开 Excel 后读取 sheet 内容
- 修改单元格后保存
- 保存时样式不变
- 数字写入后仍是数字
- `batch_update` 能处理 `D2`、`D2:D10`、搜索结果
- 没有定位到任何单元格时返回错误

## 后续计划

- 增加 JSON Schema，约束大模型 DSL 输出
- 增加上下文窗口构造器：workbook summary、sheet profile、relevant rows
- 增加操作账本：记录用户请求、DSL、结果、Diff
- 增加 Agent 包：自然语言 -> DSL -> 执行 -> 修复循环
- 增加 CLI 示例
- 支持 `.csv`
- 支持更多 Excel 样式和公式场景

## License

MIT
