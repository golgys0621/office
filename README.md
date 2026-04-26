# Office 文档处理库

**创建日期**: 2026-04-26
**版本**: v1.1
**技术栈**: Go + Node.js (Mammoth) + PDF.js + Highlight.js

---

## 一、架构设计

### 1.1 目录结构

```
OFFICE/                    # 库根目录
├── office.go              # 核心接口和注册机制
├── types.go               # 类型定义
├── docx.go                # Word文档处理 (Mammoth集成)
├── excel.go               # Excel处理 (excelize)
├── pptx.go                # PowerPoint处理 (Mammoth集成)
├── pdf.go                 # PDF处理 (PDF.js)
├── image.go               # 图片处理
├── text.go                # 文本/Markdown处理
├── csv.go                 # CSV处理
├── debug.go              # 调试日志模块
├── tracer.go              # 性能追踪模块
├── diagnostics.go          # 诊断工具模块
├── integration.go         # 集成工具
├── preview_integration.go # 预览系统集成
├── DEBUG.md              # 调试功能文档
├── README.md              # 本文档
└── scripts/
    └── mammoth/
        ├── package.json   # Node.js依赖配置
        ├── docx2html.js  # DOCX转换脚本
        └── pptx2html.js  # PPTX转换脚本
```

### 1.2 核心接口

```go
// DocumentHandler 文档处理器接口
type DocumentHandlerInterface interface {
    Name() string                    // 处理器名称
    Extensions() []string           // 支持的扩展名
    CanHandle(ext string) bool      // 是否支持该格式
    ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error)
    GetMetadata(filePath string) (*Metadata, error)
}
```

### 1.3 支持的格式

#### 📊 Excel表格
| 扩展名 | 说明 | 处理器 |
|--------|------|--------|
| .xlsx | Excel 2007+ | excelize |
| .xls | Excel 97-2003 | excelize |
| .xlsm | Excel 宏文件 | excelize |
| .et | WPS表格 | excelize |
| .ett | WPS表格模板 | excelize |
| .csv | CSV文件 | 内置 |
| .tsv | TSV文件 | 内置 |

#### 📄 Word文档
| 扩展名 | 说明 | 处理器 |
|--------|------|--------|
| .docx | Word 2007+ | Mammoth |
| .doc | Word 97-2003 | Mammoth |
| .wps | WPS文字 | Mammoth |
| .docm | Word 宏文件 | Mammoth |
| .dotx | Word模板 | Mammoth |
| .dot | Word 97模板 | Mammoth |

#### 📽️ PowerPoint演示
| 扩展名 | 说明 | 处理器 |
|--------|------|--------|
| .pptx | PowerPoint 2007+ | Mammoth |
| .ppt | PowerPoint 97-2003 | Mammoth |
| .dps | WPS演示 | Mammoth |
| .pptm | PowerPoint宏文件 | Mammoth |

#### 📋 PDF文档
| 扩展名 | 说明 | 处理器 |
|--------|------|--------|
| .pdf | PDF文档 | PDF.js |

#### 🖼️ 图片
| 扩展名 | 说明 | 处理器 |
|--------|------|--------|
| .jpg/.jpeg | JPEG图片 | 内置 |
| .png | PNG图片 | 内置 |
| .gif | GIF动画 | 内置 |
| .bmp | 位图 | 内置 |
| .webp | WebP图片 | 内置 |
| .svg | 矢量图 | 内置 |
| .ico | 图标 | 内置 |

#### 📝 文本文件
| 扩展名 | 说明 | 处理器 |
|--------|------|--------|
| .txt/.text | 纯文本 | 内置 |
| .md/.markdown | Markdown | 内置 |
| .log | 日志文件 | 内置 |
| .json | JSON | 内置 |
| .xml | XML | 内置 |
| .html | HTML | 内置 |
| .css | CSS | 内置 |
| .js | JavaScript | 内置 |
| .go/.java/.py/.sql | 编程语言 | 内置 |
| .yaml/.yml/.toml/.ini/.conf | 配置文件 | 内置 |

**共计支持 50+ 文件格式**

### 格式保留度对比

| 处理器 | 适用格式 | 格式保留度 | 特点 |
|--------|----------|-----------|------|
| **Mammoth** | .docx/.doc/.wps/.dot/.dotm/.pptx/.ppt/.dps/.pptm | 95%+ | 最佳格式保留，支持样式/表格/图片 |
| **excelize** | .xlsx/.xls/.xlsm/.et/.ett | 95%+ | 完整支持公式/图表/多Sheet |
| **PDF.js** | .pdf | 100% | 原生渲染，支持分页缩放 |
| **内置** | 图片/文本/CSV/TSV | 100% | 轻量快速，无需外部依赖 |

---

## 二、快速开始

### 2.1 安装依赖

```bash
# Go依赖
go get github.com/xuri/excelize/v2

# Node.js依赖 (用于Mammoth)
cd scripts/mammoth
npm install
```

### 2.2 基础使用

```go
package main

import (
    "fmt"
    "log"
    "github.com/golgys0621/office"
)

func main() {
    // 创建处理器
    handler := office.NewDocumentHandler()
    
    // 转换文件
    result, err := handler.Convert("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    fmt.Printf("转换成功! 格式: %s\n", result.Format)
    fmt.Printf("HTML长度: %d 字符\n", len(result.HTML))
    
    // 保存HTML
    // os.WriteFile("output.html", []byte(result.HTML), 0644)
}
```

### 2.3 转换选项

```go
opts := &office.ConversionOptions{
    EmbedImages:    true,              // 内嵌图片
    MaxWidth:      1000,             // 最大宽度
    Theme:         "light",          // 主题
    ShowPageBreaks: false,           // 显示分页
    FontScale:     1.0,             // 字体缩放
}

result, err := handler.Convert("file.docx", opts)
```

### 2.4 URL转换

```go
result, err := handler.ConvertURL("https://example.com/file.docx")
```

### 2.5 调试模式

```go
// 创建带调试的处理器
handler := office.NewDocumentHandlerWithOptions(true)

// 打印诊断信息
office.PrintDiagnostics()

// 诊断特定问题
fmt.Print(office.DiagnoseProblem("docx"))
```

---

## 三、集成到现有项目

### 3.1 Gin框架集成

```go
import "github.com/golgys0621/office"

func previewHandler(c *gin.Context) {
    url := c.Query("url")
    if url == "" {
        c.JSON(400, gin.H{"error": "缺少url参数"})
        return
    }
    
    handler := office.NewDocumentHandler()
    result, err := handler.ConvertURL(url)
    if err != nil {
        c.JSON(500, gin.H{"error": err.Error()})
        return
    }
    
    c.Data(200, "text/html; charset=utf-8", []byte(result.HTML))
}

// 路由
r.GET("/preview", previewHandler)
```

### 3.2 替换现有预览控制器

```go
// 在 preview.controller.go 中
import "github.com/golgys0621/office"

// 使用新的Office库
func OnlinePreviewPreview(c *gin.Context) {
    fileURL := c.Query("url")
    // ...
    
    handler := office.NewDocumentHandler()
    result, err := handler.ConvertURL(fileURL)
    if err != nil {
        c.Data(500, "text/html", []byte(office.OfficeErrorPage(err)))
        return
    }
    
    c.Data(200, "text/html; charset=utf-8", []byte(result.HTML))
}
```

---

## 四、Mammoth安装

Mammoth用于高质量的DOCX/PPTX转换：

```bash
cd scripts/mammoth
npm install mammoth
```

确保系统已安装Node.js。

### 4.2 PDF.js安装（可选）

PDF.js用于高质量的PDF渲染：

```bash
cd scripts/pdfjs
npm install
```

> 📌 **提示**: 如果不安装PDF.js，库将使用CDN方式渲染PDF（需要网络连接）。

---

## 五、API参考

### 5.1 DocumentHandler

| 方法 | 说明 |
|------|------|
| `NewDocumentHandler()` | 创建处理器实例 |
| `NewDocumentHandlerWithOptions(debug)` | 创建处理器（带调试） |
| `Convert(path, opts...)` | 转换本地文件 |
| `ConvertURL(url, opts...)` | 转换远程URL |
| `GetMetadata(path)` | 获取文档元信息 |
| `GetSupportedFormats()` | 获取支持的格式列表 |
| `EnableDebug()` | 启用调试模式 |
| `DisableDebug()` | 禁用调试模式 |
| `GetDebugInfo()` | 获取调试信息 |

### 5.2 HTMLResult

```go
type HTMLResult struct {
    HTML       string            // HTML内容
    CSS        string            // 额外CSS
    JS         string            // 额外JS
    Images     map[string][]byte // 内嵌图片
    Metadata   *Metadata         // 元信息
    PageCount  int               // 页数
    WordCount  int               // 字数
    Format     DocumentType      // 文档类型
}
```

### 5.3 Metadata

```go
type Metadata struct {
    Title       string    // 标题
    Author      string    // 作者
    FileSize    int64     // 文件大小
    PageCount   int       // 页数
    SheetCount  int       // Excel工作表数
    SlideCount  int       // 幻灯片数
    WordCount   int       // 字数
    Created     time.Time // 创建时间
    Modified    time.Time // 修改时间
}
```

### 5.4 调试API

```go
// 日志级别
office.SetLogLevel(office.DebugLevelDebug)

// 记录日志
office.Info("category", "message", nil)
office.Debug("category", "message", map[string]interface{}{"key": "value"})

// 性能追踪
tracer := office.StartTrace("operation")
tracer.SetTag("key", "value")
tracer.Finish()

// 诊断工具
office.PrintDiagnostics()
office.DiagnoseProblem("docx")
office.CollectDiagnostics()
```

---

## 六、错误处理

```go
result, err := handler.Convert("file.docx")
if err != nil {
    if officeErr, ok := err.(*office.OfficeError); ok {
        switch officeErr.Code {
        case office.ErrFileNotFound:
            fmt.Println("文件不存在")
        case office.ErrUnsupportedType:
            fmt.Println("不支持的格式")
        case office.ErrConversionFailed:
            fmt.Println("转换失败:", officeErr.Message)
        }
    }
}
```

---

## 七、示例代码

见 `examples/main.go`

调试功能详见 `DEBUG.md`
