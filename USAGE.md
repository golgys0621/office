# Office 文档处理库 - 使用指南

**创建日期**: 2026-04-26  
**版本**: v1.1  
**适用**: GOZSM 项目文件预览功能

---

## 一、场景概述

本库用于将 Office 文档（Word/Excel/PPT/PDF/图片/文本）转换为 HTML，实现浏览器在线预览。

### 典型使用场景

| 场景 | 示例 |
|------|------|
| 文件管理系统的在线预览 | 点击附件直接预览，无需下载 |
| 审批流程的附件查看 | 对账系统上传的发票、合同预览 |
| 文档库的在线阅读 | 规章制度、技术文档展示 |

---

## 二、三分钟快速上手

### 2.1 安装依赖

```bash
# Go 模块引入（项目已有依赖）
# 无需额外安装 Go 依赖

# Node.js 依赖（Mammoth，用于 DOCX/PPTX 转换）
cd scripts/mammoth
npm install
```

### 2.2 最小示例

```go
package main

import (
    "fmt"
    "github.com/golgys0621/office"
)

func main() {
    // 创建处理器
    handler := office.NewDocumentHandler()

    // 转换本地文件
    result, err := handler.Convert("合同.docx")
    if err != nil {
        panic(err)
    }

    fmt.Printf("转换成功！HTML长度: %d 字符\n", len(result.HTML))
}
```

### 2.3 Gin 框架集成

```go
package controllers

import (
    "github.com/gin-gonic/gin"
    "github.com/golgys0621/office"
)

func OnlinePreview(c *gin.Context) {
    fileURL := c.Query("url")
    if fileURL == "" {
        c.JSON(400, gin.H{"error": "缺少 url 参数"})
        return
    }

    handler := office.NewDocumentHandler()
    result, err := handler.ConvertURL(fileURL)
    if err != nil {
        // 返回友好的错误页面
        c.Data(500, "text/html", []byte(office.OfficeErrorPage(err)))
        return
    }

    // 返回 HTML
    c.Data(200, "text/html; charset=utf-8", []byte(result.HTML))
}
```

---

## 三、核心 API

### 3.1 DocumentHandler 处理器

```go
// 创建标准处理器
handler := office.NewDocumentHandler()

// 创建带调试的处理器
handler := office.NewDocumentHandlerWithOptions(true)
```

| 方法 | 说明 |
|------|------|
| `Convert(path)` | 转换本地文件 |
| `ConvertURL(url)` | 转换远程 URL 文件 |
| `Convert(path, opts)` | 带选项转换 |
| `GetMetadata(path)` | 获取文档元信息 |
| `GetSupportedFormats()` | 获取支持的格式列表 |
| `EnableDebug()` / `DisableDebug()` | 开关调试模式 |
| `GetDebugInfo()` | 获取调试信息 |

### 3.2 转换选项

```go
opts := &office.ConversionOptions{
    EmbedImages:    true,     // 内嵌图片到 HTML
    MaxWidth:       1200,     // 最大宽度 (px)
    Theme:          "light",  // 主题: light / dark
    ShowPageBreaks: false,    // 显示分页符
    FontScale:      1.0,      // 字体缩放
}

// 使用选项转换
result, err := handler.Convert("文件.docx", opts)
```

### 3.3 转换结果

```go
result, _ := handler.Convert("文件.xlsx")

// result 包含:
result.HTML       // HTML 内容
result.CSS        // 额外 CSS
result.JS         // 额外 JS
result.Images     // 内嵌图片 (map[name][]byte)
result.Metadata   // 文档元信息
result.PageCount  // 页数
result.WordCount  // 字数
result.Format     // 文档类型
```

---

## 四、格式支持速查

### 4.1 按类型查询

```go
// 检查是否支持某格式
handler.CanHandle(".docx")  // true
handler.CanHandle(".unknown") // false

// 获取支持的格式列表
formats := handler.GetSupportedFormats()
// ["xlsx", "xls", "xlsm", "et", "ett", "csv", ...]
```

### 4.2 处理器对照表

| 格式 | 处理器 | 说明 |
|------|--------|------|
| `.docx` `.doc` `.wps` `.dotx` `.dot` `.docm` | Mammoth | Word 文档 |
| `.xlsx` `.xls` `.xlsm` `.et` `.ett` | excelize | Excel 表格 |
| `.pptx` `.ppt` `.dps` `.pptm` | Mammoth | PPT 演示 |
| `.pdf` | PDF.js | PDF 文档 |
| `.csv` `.tsv` | 内置 | 表格文件 |
| `.jpg` `.png` `.gif` `.webp` `.svg` `.ico` `.bmp` | 内置 | 图片 |
| `.txt` `.md` `.json` `.xml` `.html` `.css` `.js` 等 | 内置 | 文本文件 |

### 4.3 格式保留度

| 处理器 | 格式保留度 | 适用场景 |
|--------|-----------|----------|
| **Mammoth** | 95%+ | Word/PPT 文档，样式/表格/图片保留好 |
| **excelize** | 95%+ | Excel 表格，公式/图表/多Sheet 完整支持 |
| **PDF.js** | 100% | PDF 原生渲染，分页缩放 |
| **内置** | 100% | 图片/文本/CSV，轻量快速 |

---

## 五、集成工具（Integration）

简化集成的封装：

```go
// 创建集成实例
i := office.NewIntegration()

// 转换本地文件
html, err := i.ConvertFile("合同.docx")

// 转换远程 URL
html, err := i.ConvertURL("https://example.com/文件.docx")

// 转换并保存
err := i.ConvertAndSave("合同.docx", "输出.html")

// 批量转换
files := []string{"a.docx", "b.xlsx", "c.pdf"}
results := i.BatchConvert(files, "./output/")

// 检查格式支持
if i.IsSupported(".docx") {
    // 支持
}
```

---

## 六、错误处理

```go
result, err := handler.Convert("文件.docx")
if err != nil {
    // 方式1: 获取错误信息
    fmt.Println(err.Error())

    // 方式2: 类型断言获取详细错误
    if officeErr, ok := err.(*office.OfficeError); ok {
        switch officeErr.Code {
        case office.ErrFileNotFound:
            fmt.Println("文件不存在")
        case office.ErrUnsupportedType:
            fmt.Println("不支持的文件格式")
        case office.ErrConversionFailed:
            fmt.Println("转换失败:", officeErr.Message)
        case office.ErrTimeout:
            fmt.Println("转换超时")
        case office.ErrInvalidFormat:
            fmt.Println("无效的文件格式")
        }
    }
}
```

### 错误页面

```go
// 在 Gin 中返回友好错误页
c.Data(500, "text/html", []byte(office.OfficeErrorPage(err)))
```

---

## 七、调试功能

### 7.1 启用调试

```go
// 创建时启用
handler := office.NewDocumentHandlerWithOptions(true)

// 运行时开关
handler.EnableDebug()
handler.DisableDebug()
```

### 7.2 诊断工具

```go
// 打印诊断报告
office.PrintDiagnostics()

// 诊断特定问题
fmt.Print(office.DiagnoseProblem("docx转换失败"))

// 收集完整诊断信息
info := office.CollectDiagnostics()
fmt.Printf("系统: %s/%s\n", info.SystemInfo.OS, info.SystemInfo.Arch)
fmt.Printf("Node.js: %v\n", info.NodeInfo.Installed)

// 导出诊断报告
data, _ := office.ExportDiagnostics("json")
```

### 7.3 性能追踪

```go
// 记录操作耗时
tracer := office.StartTrace("转换操作")
tracer.SetTag("file", "合同.docx")
tracer.Log("开始转换", nil)

// ... 执行转换 ...

tracer.Finish()

// 或使用函数式追踪
err := office.TraceFunc("operation-name", func() error {
    _, err := handler.Convert("文件.docx")
    return err
})
```

### 7.4 日志记录

```go
// 设置日志级别
office.SetLogLevel(office.DebugLevelDebug)

// 记录日志
office.Error("category", "错误信息", nil)
office.Warn("category", "警告信息", nil)
office.Info("category", "信息", nil)
office.Debug("category", "调试信息", nil)
office.Trace("category", "追踪信息", nil)
```

---

## 八、Web 前端集成

### 8.1 后端返回 HTML

```go
// Gin 控制器
func Preview(c *gin.Context) {
    fileURL := c.Query("url")
    
    handler := office.NewDocumentHandler()
    result, err := handler.ConvertURL(fileURL)
    
    if err != nil {
        c.JSON(500, gin.H{"error": err.Error()})
        return
    }
    
    c.Data(200, "text/html; charset=utf-8", []byte(result.HTML))
}
```

### 8.2 前端预览组件

```vue
<template>
  <view class="preview-container">
    <web-view :src="previewUrl" v-if="previewUrl"></web-view>
    <view class="loading" v-else>加载中...</view>
  </view>
</template>

<script>
export default {
  data() {
    return {
      previewUrl: ''
    }
  },
  onLoad(options) {
    const fileUrl = encodeURIComponent(options.url)
    // 指向后端预览接口
    this.previewUrl = `http://your-server/api/online-preview?url=${fileUrl}`
  }
}
</script>
```

### 8.3 注意事项

1. **跨域问题**: 确保后端配置了正确的 CORS 头
2. **URL 编码**: 文件 URL 需要编码后再传递
3. **文件大小**: 大文件转换耗时较长，建议添加超时
4. **图片路径**: Mammoth 转换的图片需使用 `EmbedImages: true`

---

## 九、常见问题

### Q1: 转换失败提示 "Node.js not found"

**原因**: Mammoth 需要 Node.js 环境

**解决**:
```bash
# 安装 Node.js
# Windows: https://nodejs.org/
# 然后安装 mammoth
cd scripts/mammoth
npm install
```

### Q2: DOCX 转换样式丢失

**原因**: Mammoth 对复杂样式支持有限

**解决**:
- 简化文档样式
- 或考虑用 PDF 作为中间格式
- 使用 `EmbedImages: true` 确保图片正确嵌入

### Q3: Excel 表格数据量很大

**解决**:
- 设置 `MaxWidth` 限制宽度
- 考虑分 Sheet 显示
- 大文件可添加进度提示

### Q4: 转换超时

**原因**: 文件过大或网络慢

**解决**:
```go
// 在 Gin 中设置超时
ctx, cancel := context.WithTimeout(c.Request.Context(), 30*time.Second)
defer cancel()

// 使用带超时的转换
result, err := handler.ConvertURL(fileURL)
```

### Q5: 不支持的文件格式

**提示**: 返回友好的错误信息
```go
if !handler.CanHandle(ext) {
    c.JSON(400, gin.H{"error": "不支持的文件格式: " + ext})
    return
}
```

---

## 十、最佳实践

### 10.1 推荐配置

```go
handler := office.NewDocumentHandlerWithOptions(true)

// 设置 Mammoth 路径（如果非标准位置）
handler.SetMammothPath("scripts/mammoth/node_modules/.bin/mammoth")
```

### 10.2 缓存策略

```go
// 对于频繁访问的文件，考虑缓存转换结果
// 缓存 key: md5(url + file_mtime)
// 缓存时间: 根据业务需求设置
```

### 10.3 监控建议

```go
// 记录转换耗时
tracer := office.StartTrace("preview")
defer tracer.Finish()

result, err := handler.ConvertURL(url)
if err != nil {
    tracer.SetTag("status", "failed")
    // 上报到监控系统
}
```

### 10.4 权限检查

```go
func OnlinePreview(c *gin.Context) {
    // 1. 检查用户权限
    if !checkPermission(c) {
        c.JSON(403, gin.H{"error": "无权限"})
        return
    }
    
    // 2. 检查文件存在
    url := c.Query("url")
    if !isAccessible(url) {
        c.JSON(403, gin.H{"error": "文件不可访问"})
        return
    }
    
    // 3. 执行转换
    // ...
}
```

---

## 十一、路由配置示例

```go
// routers/api.router.go

func SetupRouter() *gin.Engine {
    r := gin.Default()
    
    // 在线预览接口
    r.GET("/api/online-preview", controllers.OnlinePreview)
    
    // 或作为中间件
    r.GET("/file/preview", 
        middleware.CheckAuth(),
        controllers.OnlinePreview,
    )
    
    return r
}
```

---

## 十二、完整示例

```go
package main

import (
    "fmt"
    "log"
    "net/http"
    "time"

    "github.com/gin-gonic/gin"
    "github.com/golgys0621/office"
)

func main() {
    r := gin.Default()
    
    // 启用调试模式
    office.SetLogLevel(office.DebugLevelInfo)
    
    r.GET("/preview", func(c *gin.Context) {
        url := c.Query("url")
        if url == "" {
            c.JSON(400, gin.H{"error": "缺少 url 参数"})
            return
        }
        
        handler := office.NewDocumentHandlerWithOptions(true)
        
        // 带超时
        start := time.Now()
        result, err := handler.ConvertURL(url)
        if err != nil {
            log.Printf("预览失败: %v", err)
            c.Data(500, "text/html", []byte(office.OfficeErrorPage(err)))
            return
        }
        
        fmt.Printf("转换耗时: %v\n", time.Since(start))
        
        c.Data(200, "text/html; charset=utf-8", []byte(result.HTML))
    })
    
    log.Println("预览服务启动: http://localhost:8080")
    r.Run(":8080")
}
```

---

## 附录: 相关文件

| 文件 | 说明 |
|------|------|
| `README.md` | 完整技术文档 |
| `DEBUG.md` | 调试功能详解 |
| `examples/main.go` | 完整示例代码 |
| `types.go` | 类型定义 |
| `integration.go` | 集成工具源码 |
