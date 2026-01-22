# Impex

Impex 是一个基于 [excelize](https://github.com/xuri/excelize) 构建的轻量级、泛型支持的 Go 语言 Excel 导入导出库。它旨在通过结构体标签（Struct Tags）和泛型（Generics）极大简化 Excel 处理流程，减少样板代码。

## 特性 (Features)

- **泛型支持 (Generics)**: 直接返回 `[]T` 类型切片，告别繁琐的类型断言。
- **零配置 (Zero Boilerplate)**: 使用 `excel:"列名"` 标签即可完成映射配置。
- **动态列支持 (Dynamic Columns)**: 支持 `excel:"extra"` 或 `excel:"*"` 自动捕获未定义的不定长列（如时间序列数据）。
- **增强导出 (Enhanced Export)**: 支持设置列宽 (`width:20`)、强制文本格式 (`text`) 以及自定义转换器。

## 安装 (Installation)

```bash
go get github.com/enterShuIoT/impex
```

## 快速上手 (Quick Start)

### 1. 导入 (Import)

#### 基础导入

```go
package main

import (
    "fmt"
    "github.com/enterShuIoT/impex/importer"
)

type User struct {
    Name  string `excel:"姓名"`
    Age   int    `excel:"年龄"`
    Phone string `excel:"手机号"`
}

func main() {
    // 1. 创建配置 (最简配置)
    config := &importer.ExcelImportConfig[User]{
        SheetName: "Sheet1",
    }

    // 2. 创建导入器
    imp := importer.NewExcelImporter(config)

    // 3. 执行导入
    users, err := imp.ImportLocal("users.xlsx") // 返回 []User
    if err != nil {
        panic(err)
    }

    for _, u := range users {
        fmt.Printf("Name: %s, Age: %d\n", u.Name, u.Age)
    }
}
```

#### 动态列导入 (例如：时间序列数据)

处理像 "00:30", "01:00" 这种不固定的列名时，可以使用 `excel:"extra"` 将其捕获到一个 `map` 中。

```go
type LoadForecast struct {
    ClientAccount string            `excel:"用户编号"`
    Date          string            `excel:"日期"`
    // 自动捕获所有未映射的列到 TimeData map 中
    TimeData      map[string]string `excel:"extra"` 
}

func main() {
    imp := importer.NewExcelImporter(&importer.ExcelImportConfig[LoadForecast]{})
    data, _ := imp.ImportLocal("forecast.xlsx")
    
    // data[0].TimeData["00:30"] 即可获取对应值
}
```

### 2. 导出 (Export)

#### 基础导出与格式控制

支持在 Tag 中直接定义导出样式：
- `text`: 强制该列为文本格式（防止长数字变成科学计数法）。
- `width:N`: 设置列宽。

```go
package main

import (
    "github.com/enterShuIoT/impex/exporter"
)

type ExportItem struct {
    Name      string   `excel:"名称,width:20"`    // 设置列宽为 20
    Account   string   `excel:"账号,text"`        // 强制文本格式
    Score     float64  `excel:"分数"`
    Reference *float64 `excel:"参考值"`            // 自动处理指针，nil 输出为空
}

func main() {
    data := []ExportItem{
        {Name: "张三", Account: "123456789012345", Score: 98.5},
    }

    // 1. 创建配置
    config := &exporter.ExcelExportConfig[ExportItem]{
        FileName: "report.xlsx",
    }

    // 2. 创建导出器
    exp := exporter.NewExcelExporter(config)

    // 3. 执行导出
    resp, err := exp.Export(data)
    if err != nil {
        panic(err)
    }

    // resp.Content 包含了 Excel 文件的 []byte 内容，可直接写入文件或通过 HTTP 下载
}
```

#### 自定义转换器 (Custom Converters)

```go
config := &exporter.ExcelExportConfig[MyData]{
    CustomConverters: map[string]func(any) any{
        "Status": func(v any) any {
            // 将状态码转换为中文
            if v.(int) == 1 { return "启用" }
            return "禁用"
        },
        "Value": func(v any) any {
            // 保留4位小数
            val := v.(float64)
            return math.Round(val*10000) / 10000
        },
    },
}
```
