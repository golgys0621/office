# Office文档库调试功能文档

## 概述

为Office文档库添加了完整的调试和诊断功能，包括日志记录、性能追踪、问题诊断等。

## 新增模块

### 1. debug.go - 调试日志模块

提供结构化日志记录功能。

#### 主要功能

- **日志级别控制**: OFF → ERROR → WARN → INFO → DEBUG → TRACE
- **多输出方式**: 控制台(支持彩色)/文件/JSON格式
- **日志缓冲**: 内存缓冲+文件持久化
- **性能记录**: 自动记录操作耗时

#### API

```go
// 创建日志器
logger := office.NewLogger(&office.DebugConfig{
    Level:         office.DebugLevelDebug,
    EnableConsole: true,
    EnableFile:    true,
    LogFilePath:   "logs/office.log",
    EnableColors:  true,
})

// 记录日志
logger.Error("category", "message", nil)
logger.Warn("category", "message", nil)
logger.Info("category", "message", nil)
logger.Debug("category", "message", nil)
logger.Trace("category", "message", nil)

// 性能记录
logger.RecordPerf("operation_name", time.Duration)

// 获取性能统计
stats := logger.GetAllPerfStats()

// 获取日志条目
entries := logger.GetEntries(office.DebugLevelError, "category", 50)
```

#### 全局函数

```go
// 设置日志级别
office.SetLogLevel(office.DebugLevelDebug)

// 便捷日志函数
office.Error("test", "错误信息", nil)
office.Warn("test", "警告信息", nil)
office.Info("test", "信息", nil)
office.Debug("test", "调试信息", nil)
office.Trace("test", "追踪信息", nil)
```

---

### 2. tracer.go - 性能追踪模块

提供分布式追踪风格的操作追踪功能。

#### 主要功能

- **Span追踪**: 支持命名span和嵌套关系
- **标签和日志**: 每个span可附加键值对标签和日志
- **性能统计**: 自动统计调用次数、总耗时、最小/最大/平均耗时
- **错误追踪**: 记录span执行中的错误

#### API

```go
// 开始追踪
span := office.StartSpan("operation_name", 
    office.WithTag("key1", "value1"),
    office.WithTag("key2", 123),
)

// 添加日志
span.Log("处理步骤1完成", map[string]interface{}{"count": 10})

// 添加标签
span.SetTag("result", "success")

// 结束span
office.StopSpan(span, nil) // 成功
office.StopSpan(span, err) // 失败
```

#### TraceOperation 便捷方式

```go
// 创建追踪操作
op := office.StartTrace("my-operation")

// 添加标签
op.SetTag("param1", "value1")

// 记录日志
op.Log("开始处理", nil)

// 结束
op.End(nil)  // 无错误
op.Finish()  // 同上
```

#### 追踪函数执行

```go
// 自动追踪函数
err := office.TraceFunc("file-operation", func() error {
    // 你的代码
    return nil
})

// 获取结果和追踪
result, err := office.TraceFuncResult("my-func", func() (interface{}, error) {
    return doSomething()
})
```

#### 获取统计

```go
stats := office.GetTracerStats()
// stats.TotalSpans   - 总跨度数
// stats.ActiveSpans - 活跃跨度数
// stats.ErrorCount  - 错误数
// stats.TotalDuration - 总耗时
```

---

### 3. diagnostics.go - 诊断工具模块

提供完整的环境诊断和问题排查功能。

#### 主要功能

- **系统信息**: OS、CPU、内存、主机名等
- **Go环境**: Go版本、平台信息
- **Node.js环境**: 安装状态、版本、路径
- **依赖检查**: mammoth、pdf.js等
- **错误追踪**: 最近错误日志
- **性能统计**: 操作耗时统计
- **问题诊断**: 针对特定问题的诊断和建议

#### API

```go
// 收集完整诊断信息
info := office.CollectDiagnostics()

// 诊断特定问题
report := office.DiagnoseProblem("docx") // "excel", "pptx", "pdf", "node"
fmt.Print(report)

// 检查环境
ok, errors := office.CheckEnvironment()
if !ok {
    for _, e := range errors {
        fmt.Println(e)
    }
}

// 打印诊断报告
office.PrintDiagnostics()

// 导出诊断报告
data, _ := office.ExportDiagnostics("json")  // "json" 或 "text"
fmt.Println(string(data))

// 打印版本信息
office.PrintVersionInfo()
```

#### DiagnosticInfo 结构

```go
type DiagnosticInfo struct {
    SystemInfo    SystemInfo           // 系统信息
    GoInfo        GoInfo              // Go环境
    NodeInfo      NodeInfo            // Node.js环境
    Environment   map[string]string   // 相关环境变量
    Dependencies  []DependencyInfo    // 依赖状态
    RecentErrors  []LogEntry          // 最近错误
    PerfStats     []*PerfStats        // 性能统计
    TracerStats   TracerStats         // 追踪统计
}
```

---

### 4. DocumentHandler调试集成

处理器现已集成调试功能。

#### API

```go
// 创建带调试的处理器
handler := office.NewDocumentHandlerWithOptions(true)

// 启用/禁用调试
handler.EnableDebug()
handler.DisableDebug()

// 检查状态
if handler.IsDebugEnabled() {
    fmt.Println("调试模式已启用")
}

// 获取调试信息
debugInfo := handler.GetDebugInfo()
// debugInfo.DebugEnabled - 是否启用
// debugInfo.TracerStats  - 追踪统计
// debugInfo.PerfStats    - 性能统计
// debugInfo.ActiveSpans  - 活跃span
// debugInfo.RecentLogs  - 最近日志
```

#### 调试输出示例

```
[2026-04-26 10:00:00.123] INFO [convert] Starting conversion {"file": "test.docx", "format": "docx"}
[2026-04-26 10:00:00.456] DEBUG [convert] Conversion completed {"file": "test.docx", "handler": "docx", "duration": "333ms", "success": true}
```

---

## 使用示例

### 完整示例

```go
package main

import (
    "fmt"
    "github.com/golgys0621/office"
)

func main() {
    // 1. 设置日志级别
    office.SetLogLevel(office.DebugLevelDebug)
    
    // 2. 启用全局追踪
    office.EnableTracer()
    
    // 3. 创建带调试的处理器
    handler := office.NewDocumentHandlerWithOptions(true)
    
    // 4. 执行转换（自动记录日志和性能）
    result, err := handler.Convert("document.docx")
    if err != nil {
        fmt.Printf("转换失败: %v\n", err)
    }
    
    // 5. 打印诊断报告
    office.PrintDiagnostics()
    
    // 6. 导出诊断报告
    data, _ := office.ExportDiagnostics("json")
    fmt.Println(string(data))
}
```

### 问题排查示例

```go
// 当遇到问题时
func diagnoseAndFix() {
    // 收集诊断信息
    info := office.CollectDiagnostics()
    
    // 检查Node.js
    if !info.NodeInfo.Installed {
        fmt.Println("需要安装Node.js")
        return
    }
    
    // 检查依赖
    for _, dep := range info.Dependencies {
        if dep.Status != "ok" {
            fmt.Printf("%s 未安装或版本不匹配\n", dep.Name)
        }
    }
    
    // 查看最近错误
    for _, err := range info.RecentErrors {
        fmt.Printf("[%s] %s: %s\n", err.Time, err.Category, err.Message)
    }
}
```

---

## 环境变量

```bash
# 设置日志级别
export OFFICE_DEBUG=debug  # off/error/warn/info/debug/trace

# 设置Node环境
export NODE_ENV=production
```

---

## 常见问题排查

### DOCX转换失败

1. 检查Node.js是否安装: `node --version`
2. 检查Mammoth是否安装: `cd scripts/mammoth && npm install`
3. 运行诊断: `office.DiagnoseProblem("docx")`

### 性能问题

1. 查看性能统计: `handler.GetDebugInfo().PerfStats`
2. 启用追踪: `office.EnableTracer()`
3. 分析慢操作: 关注Avg耗时最长的操作

### 环境问题

1. 运行环境检查: `office.CheckEnvironment()`
2. 打印完整诊断: `office.PrintDiagnostics()`
3. 导出JSON报告: `office.ExportDiagnostics("json")`
