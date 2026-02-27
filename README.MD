# LegacyOfficeReader

[![Go Reference](https://pkg.go.dev/badge/github.com/shakinm/xlsReader.svg)](https://pkg.go.dev/github.com/shakinm/xlsReader)

A Go library for reading and converting legacy Microsoft Office binary file formats (XLS, DOC, PPT) to modern Office Open XML formats (XLSX, DOCX, PPTX).

**English** | [中文](#中文文档)

## Features

- **XLS Reader**: Parse Excel 97-2003 binary files (.xls)
- **DOC Reader**: Parse Word 97-2003 binary files (.doc)
- **PPT Reader**: Parse PowerPoint 97-2003 binary files (.ppt)
- **Format Conversion**: Convert legacy formats to modern Office Open XML formats
  - XLS → XLSX
  - DOC → DOCX
  - PPT → PPTX

## Installation

```bash
go get github.com/shakinm/xlsReader
```

## Usage

### Reading XLS Files

```go
package main

import (
    "fmt"
    "github.com/shakinm/xlsReader/xls"
    "log"
)

func main() {
    workbook, err := xls.OpenFile("spreadsheet.xls")
    if err != nil {
        log.Fatal(err)
    }

    // Get number of sheets
    fmt.Printf("Sheets: %d\n", workbook.GetNumberSheets())

    // Get first sheet
    sheet, err := workbook.GetSheet(0)
    if err != nil {
        log.Fatal(err)
    }

    fmt.Printf("Sheet name: %s\n", sheet.GetName())
    fmt.Printf("Rows: %d\n", sheet.GetNumberRows())

    // Iterate through rows and columns
    for i := 0; i <= sheet.GetNumberRows(); i++ {
        if row, err := sheet.GetRow(i); err == nil {
            for j := 0; j < len(row.GetCols()); j++ {
                if cell, err := row.GetCol(j); err == nil {
                    fmt.Printf("Cell [%d,%d]: %s\n", i, j, cell.GetString())
                }
            }
        }
    }
}
```

### Reading DOC Files

```go
package main

import (
    "fmt"
    "github.com/shakinm/xlsReader/doc"
    "log"
)

func main() {
    document, err := doc.OpenFile("document.doc")
    if err != nil {
        log.Fatal(err)
    }

    // Get plain text
    fmt.Println(document.GetText())

    // Get images
    images := document.GetImages()
    for i, img := range images {
        fmt.Printf("Image %d: format=%s, size=%d bytes\n", i, img.Format, len(img.Data))
    }
}
```

### Reading PPT Files

```go
package main

import (
    "fmt"
    "github.com/shakinm/xlsReader/ppt"
    "log"
)

func main() {
    presentation, err := ppt.OpenFile("presentation.ppt")
    if err != nil {
        log.Fatal(err)
    }

    // Get slide text
    for i, slide := range presentation.GetSlides() {
        fmt.Printf("Slide %d: %s\n", i+1, slide.GetText())
    }
}
```

### Converting XLS to XLSX

```go
package main

import (
    "github.com/shakinm/xlsReader/convert/xlsconv"
    "log"
)

func main() {
    err := xlsconv.ConvertFile("input.xls", "output.xlsx")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Converting DOC to DOCX

```go
package main

import (
    "github.com/shakinm/xlsReader/convert/docconv"
    "log"
)

func main() {
    err := docconv.ConvertFile("input.doc", "output.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Converting PPT to PPTX

```go
package main

import (
    "github.com/shakinm/xlsReader/convert/pptconv"
    "log"
)

func main() {
    err := pptconv.ConvertFile("input.ppt", "output.pptx")
    if err != nil {
        log.Fatal(err)
    }
}
```

## Project Structure

```
LegacyOfficeReader/
├── cfb/                    # Compound File Binary format parser
├── xls/                    # XLS file parser
│   ├── record/            # BIFF record types
│   └── structure/         # Data structures
├── doc/                    # DOC file parser
├── ppt/                    # PPT file parser
├── common/                 # Shared utilities (image handling)
├── helpers/                # Helper functions (date, encoding)
├── convert/                # Format converters
│   ├── xlsconv/           # XLS → XLSX conversion
│   ├── docconv/           # DOC → DOCX conversion
│   └── pptconv/           # PPT → PPTX conversion
└── cmd/                    # Command-line tools
    └── convert/           # Batch conversion tool
```

## Supported Features

### XLS/XLSX
- Cell values (strings, numbers, formulas)
- Multiple sheets
- Cell formatting
- Date/number formatting

### DOC/DOCX
- Plain text extraction
- Paragraph formatting
- Image extraction
- Headers and footers
- Text boxes

### PPT/PPTX
- Slide text extraction
- Image extraction
- Slide content preservation

## Dependencies

- [github.com/metakeule/fmtdate](https://github.com/metakeule/fmtdate) - Date formatting
- [golang.org/x/text](https://golang.org/x/text) - Text encoding support

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

---

# 中文文档

一个用于读取和转换旧版 Microsoft Office 二进制文件格式（XLS, DOC, PPT）到现代 Office Open XML 格式（XLSX, DOCX, PPTX）的 Go 语言库。

## 功能特性

- **XLS 读取器**：解析 Excel 97-2003 二进制文件（.xls）
- **DOC 读取器**：解析 Word 97-2003 二进制文件（.doc）
- **PPT 读取器**：解析 PowerPoint 97-2003 二进制文件（.ppt）
- **格式转换**：将旧格式转换为现代 Office Open XML 格式
  - XLS → XLSX
  - DOC → DOCX
  - PPT → PPTX

## 安装

```bash
go get github.com/shakinm/xlsReader
```

## 使用示例

### 读取 XLS 文件

```go
package main

import (
    "fmt"
    "github.com/shakinm/xlsReader/xls"
    "log"
)

func main() {
    workbook, err := xls.OpenFile("spreadsheet.xls")
    if err != nil {
        log.Fatal(err)
    }

    // 获取工作表数量
    fmt.Printf("工作表数量: %d\n", workbook.GetNumberSheets())

    // 获取第一个工作表
    sheet, err := workbook.GetSheet(0)
    if err != nil {
        log.Fatal(err)
    }

    fmt.Printf("工作表名称: %s\n", sheet.GetName())
    fmt.Printf("行数: %d\n", sheet.GetNumberRows())

    // 遍历行和列
    for i := 0; i <= sheet.GetNumberRows(); i++ {
        if row, err := sheet.GetRow(i); err == nil {
            for j := 0; j < len(row.GetCols()); j++ {
                if cell, err := row.GetCol(j); err == nil {
                    fmt.Printf("单元格 [%d,%d]: %s\n", i, j, cell.GetString())
                }
            }
        }
    }
}
```

### 格式转换

```go
// XLS 转 XLSX
xlsconv.ConvertFile("input.xls", "output.xlsx")

// DOC 转 DOCX
docconv.ConvertFile("input.doc", "output.docx")

// PPT 转 PPTX
pptconv.ConvertFile("input.ppt", "output.pptx")
```

## 许可证

本项目采用 GNU General Public License v3.0 许可证 - 详见 [LICENSE](LICENSE) 文件。
 
