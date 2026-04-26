package main

import (
	"fmt"
	"log"
	"os"

	"github.com/golgys0621/office"
)

func main() {
	// 创建Office处理器（启用调试模式）
	handler := office.NewDocumentHandlerWithOptions(true)

	// 获取支持的格式
	formats := handler.GetSupportedFormats()
	fmt.Println("支持的格式:", formats)
	fmt.Println()

	// 演示各种转换
	demo(handler)

	// 打印诊断信息
	fmt.Println("\n=== 诊断信息 ===")
	office.PrintDiagnostics()
}

func demo(handler *office.DocumentHandler) {
	// 示例1: 转换DOCX文件
	convertFile(handler, "document.docx")

	// 示例2: 转换Excel文件
	convertFile(handler, "spreadsheet.xlsx")

	// 示例3: 转换PPTX文件
	convertFile(handler, "presentation.pptx")

	// 示例4: 转换PDF文件
	convertFile(handler, "document.pdf")

	// 示例5: 转换图片
	convertFile(handler, "image.png")

	// 示例6: 转换文本文件
	convertFile(handler, "readme.md")

	// 示例7: 转换CSV
	convertFile(handler, "data.csv")

	// 示例8: 批量转换
	demoBatchConversion(handler)

	// 示例9: 使用诊断工具
	demoDiagnostics()
}

func convertFile(handler *office.DocumentHandler, filePath string) {
	// 检查文件是否存在
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		fmt.Printf("文件不存在: %s\n", filePath)
		return
	}

	// 转换文件
	result, err := handler.Convert(filePath)
	if err != nil {
		log.Printf("转换失败 %s: %v\n", filePath, err)
		return
	}

	// 输出结果
	fmt.Printf("转换成功: %s\n", filePath)
	fmt.Printf("  格式: %s\n", result.Format)
	fmt.Printf("  HTML长度: %d 字符\n", len(result.HTML))

	// 如果有元信息
	if result.Metadata != nil {
		fmt.Printf("  文件大小: %d bytes\n", result.Metadata.FileSize)
	}

	// 保存HTML结果（可选）
	outputPath := filePath + ".html"
	if err := os.WriteFile(outputPath, []byte(result.HTML), 0644); err == nil {
		fmt.Printf("  已保存到: %s\n", outputPath)
	}
	fmt.Println()
}

// 示例: 从URL转换
func demoURLConversion(handler *office.DocumentHandler) {
	url := "https://example.com/document.docx"

	result, err := handler.ConvertURL(url)
	if err != nil {
		log.Printf("URL转换失败: %v\n", err)
		return
	}

	fmt.Printf("URL转换成功!\n")
	fmt.Printf("  格式: %s\n", result.Format)
	fmt.Printf("  HTML长度: %d 字符\n", len(result.HTML))
}

// 示例: 批量转换
func demoBatchConversion(handler *office.DocumentHandler) {
	fmt.Println("=== 批量转换示例 ===")

	files := []string{
		"doc1.docx",
		"doc2.xlsx",
		"doc3.pptx",
	}

	for _, file := range files {
		if _, err := os.Stat(file); os.IsNotExist(err) {
			fmt.Printf("跳过 (不存在): %s\n", file)
			continue
		}

		// 使用追踪器记录
		tracer := office.StartTrace("batch:" + file)
		tracer.SetTag("file", file)

		result, err := handler.Convert(file)
		if err != nil {
			tracer.End(err)
			fmt.Printf("转换失败: %s - %v\n", file, err)
			continue
		}

		tracer.Finish()
		fmt.Printf("批量转换: %s -> %s (%d 字符)\n", file, result.Format, len(result.HTML))
	}
	fmt.Println()
}

// 示例: 诊断工具使用
func demoDiagnostics() {
	fmt.Println("=== 诊断示例 ===")

	// 收集诊断信息
	info := office.CollectDiagnostics()
	fmt.Printf("系统: %s/%s, CPU: %d\n", info.SystemInfo.OS, info.SystemInfo.Arch, info.SystemInfo.NumCPU)
	fmt.Printf("Node.js: %v\n", info.NodeInfo.Installed)

	// 诊断特定问题
	fmt.Println("\n诊断DOCX问题:")
	fmt.Print(office.DiagnoseProblem("docx转换失败"))

	// 导出诊断报告
	if data, err := office.ExportDiagnostics("json"); err == nil {
		fmt.Println("\n诊断报告(JSON):", string(data[:min(200, len(data))])+"...")
	}
}

// 示例: 调试信息获取
func demoDebugInfo() {
	handler := office.NewDocumentHandlerWithOptions(true)

	// 执行一些转换
	handler.Convert("test.docx")

	// 获取调试信息
	info := handler.GetDebugInfo()
	fmt.Printf("调试状态: %v\n", info.DebugEnabled)
	fmt.Printf("追踪统计: %+v\n", info.TracerStats)
	fmt.Printf("性能统计: %d 项\n", len(info.PerfStats))
}

// 示例: 使用日志级别
func demoLogLevels() {
	// 设置日志级别
	office.SetLogLevel(office.DebugLevelDebug)

	// 记录不同级别的日志
	office.Error("test", "这是一个错误", nil)
	office.Warn("test", "这是一个警告", nil)
	office.Info("test", "这是一个信息", nil)
	office.Debug("test", "这是一个调试信息", nil)
	office.Trace("test", "这是一个追踪信息", nil)

	// 性能追踪示例
	office.StartTrace("test-operation").End(nil)
}

// 示例: 使用TraceFunc
func demoTraceFunc() {
	// 自动追踪函数执行
	err := office.TraceFunc("file-operation", func() error {
		// 模拟一些操作
		handler := office.NewDocumentHandler()
		_, _ = handler.Convert("test.docx")
		return nil
	})

	if err != nil {
		fmt.Printf("操作失败: %v\n", err)
	}
}

// 示例: 获取依赖版本信息
func demoVersionInfo() {
	office.PrintVersionInfo()
}
