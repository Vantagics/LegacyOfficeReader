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

    // Get formatted content (paragraphs, styles, etc.)
    formatted := document.GetFormattedContent()
    if formatted != nil {
        for i, para := range formatted.Paragraphs {
            fmt.Printf("Paragraph %d: %s\n", i, para.Text)
            // Access character formatting
            for _, run := range para.Runs {
                if run.Bold {
                    fmt.Printf("  Bold text: %s\n", run.Text)
                }
            }
        }

        // Get headers and footers
        for _, header := range formatted.Headers {
            fmt.Printf("Header: %s\n", header.Text)
        }
        for _, footer := range formatted.Footers {
            fmt.Printf("Footer: %s\n", footer.Text)
        }
    }

    // Get fonts used in document
    fonts := document.GetFonts()
    fmt.Printf("Fonts: %v\n", fonts)

    // Get styles
    styles := document.GetStyles()
    fmt.Printf("Styles: %v\n", styles)
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

    // Get number of slides
    fmt.Printf("Total slides: %d\n", presentation.GetNumberSlides())

    // Get slide dimensions (in EMU units)
    width, height := presentation.GetSlideSize()
    fmt.Printf("Slide size: %d x %d EMU\n", width, height)

    // Iterate through slides
    for i := 0; i < presentation.GetNumberSlides(); i++ {
        slide, err := presentation.GetSlide(i)
        if err != nil {
            log.Fatal(err)
        }

        fmt.Printf("\n--- Slide %d ---\n", i+1)

        // Get all text blocks
        texts := slide.GetTexts()
        for j, text := range texts {
            fmt.Printf("Text %d: %s\n", j, text)
        }

        // Get shapes with formatting info
        shapes := slide.GetShapes()
        for j, shape := range shapes {
            fmt.Printf("Shape %d: type=%d, left=%d, top=%d, width=%d, height=%d\n",
                j, shape.ShapeType, shape.Left, shape.Top, shape.Width, shape.Height)
        }
    }

    // Get all images from presentation
    images := presentation.GetImages()
    for i, img := range images {
        fmt.Printf("Image %d: format=%s, size=%d bytes\n", i, img.Format, len(img.Data))
    }

    // Get fonts used in presentation
    fonts := presentation.GetFonts()
    fmt.Printf("Fonts: %v\n", fonts)
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

### 读取 DOC 文件

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

    // 获取纯文本内容
    fmt.Println(document.GetText())

    // 获取图片
    images := document.GetImages()
    for i, img := range images {
        fmt.Printf("图片 %d: 格式=%s, 大小=%d 字节\n", i, img.Format, len(img.Data))
    }

    // 获取格式化内容（段落、样式等）
    formatted := document.GetFormattedContent()
    if formatted != nil {
        for i, para := range formatted.Paragraphs {
            fmt.Printf("段落 %d: %s\n", i, para.Text)
            // 访问字符格式
            for _, run := range para.Runs {
                if run.Bold {
                    fmt.Printf("  粗体文本: %s\n", run.Text)
                }
            }
        }

        // 获取页眉页脚
        for _, header := range formatted.Headers {
            fmt.Printf("页眉: %s\n", header.Text)
        }
        for _, footer := range formatted.Footers {
            fmt.Printf("页脚: %s\n", footer.Text)
        }
    }

    // 获取文档中使用的字体
    fonts := document.GetFonts()
    fmt.Printf("字体: %v\n", fonts)

    // 获取样式
    styles := document.GetStyles()
    fmt.Printf("样式: %v\n", styles)
}
```

### 读取 PPT 文件

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

    // 获取幻灯片数量
    fmt.Printf("幻灯片总数: %d\n", presentation.GetNumberSlides())

    // 获取幻灯片尺寸（EMU单位）
    width, height := presentation.GetSlideSize()
    fmt.Printf("幻灯片尺寸: %d x %d EMU\n", width, height)

    // 遍历幻灯片
    for i := 0; i < presentation.GetNumberSlides(); i++ {
        slide, err := presentation.GetSlide(i)
        if err != nil {
            log.Fatal(err)
        }

        fmt.Printf("\n--- 幻灯片 %d ---\n", i+1)

        // 获取所有文本块
        texts := slide.GetTexts()
        for j, text := range texts {
            fmt.Printf("文本 %d: %s\n", j, text)
        }

        // 获取形状及其格式信息
        shapes := slide.GetShapes()
        for j, shape := range shapes {
            fmt.Printf("形状 %d: 类型=%d, 左=%d, 上=%d, 宽=%d, 高=%d\n",
                j, shape.ShapeType, shape.Left, shape.Top, shape.Width, shape.Height)
        }
    }

    // 获取演示文稿中的所有图片
    images := presentation.GetImages()
    for i, img := range images {
        fmt.Printf("图片 %d: 格式=%s, 大小=%d 字节\n", i, img.Format, len(img.Data))
    }

    // 获取演示文稿中使用的字体
    fonts := presentation.GetFonts()
    fmt.Printf("字体: %v\n", fonts)
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
 
