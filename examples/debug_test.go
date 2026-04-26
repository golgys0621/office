package main

import (
	"fmt"

	"github.com/golgys0621/office"
)

func main() {
	fmt.Println("=== Office文档库调试功能测试 ===\n")

	// 1. 测试诊断信息收集
	fmt.Println("1. 测试诊断信息收集:")
	info := office.CollectDiagnostics()
	fmt.Printf("   系统: %s/%s\n", info.SystemInfo.OS, info.SystemInfo.Arch)
	fmt.Printf("   Node.js已安装: %v\n", info.NodeInfo.Installed)
	fmt.Println()

	// 2. 测试版本信息
	fmt.Println("2. 测试版本信息:")
	office.PrintVersionInfo()
	fmt.Println()

	// 3. 测试调试日志
	fmt.Println("3. 测试调试日志:")
	office.SetLogLevel(office.DebugLevelDebug)
	office.Error("test", "这是一个错误信息", map[string]interface{}{"key": "value"})
	office.Warn("test", "这是一个警告信息", nil)
	office.Info("test", "这是一个信息", nil)
	office.Debug("test", "这是一个调试信息", nil)
	fmt.Println()

	// 4. 测试性能追踪
	fmt.Println("4. 测试性能追踪:")
	office.EnableTracer()
	tracer := office.StartTrace("test-operation")
	tracer.SetTag("param1", "value1")
	tracer.Log("开始处理", nil)
	tracer.Finish()
	stats := office.GetTracerStats()
	fmt.Printf("   追踪统计: %+v\n", stats)
	fmt.Println()

	// 5. 测试问题诊断
	fmt.Println("5. 测试问题诊断:")
	fmt.Print(office.DiagnoseProblem("docx"))
	fmt.Println()

	// 6. 测试环境检查
	fmt.Println("6. 测试环境检查:")
	ok, errors := office.CheckEnvironment()
	fmt.Printf("   环境检查: %v\n", ok)
	if !ok {
		for _, e := range errors {
			fmt.Printf("   - %s\n", e)
		}
	}
	fmt.Println()

	// 7. 测试处理器调试模式
	fmt.Println("7. 测试处理器调试模式:")
	handler := office.NewDocumentHandlerWithOptions(true)
	fmt.Printf("   调试模式已启用: %v\n", handler.IsDebugEnabled())
	fmt.Println()

	fmt.Println("=== 测试完成 ===")
}
