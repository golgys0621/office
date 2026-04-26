package office

import (
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"
)

// DocumentHandler 文档处理器
type DocumentHandler struct {
	handlers    map[string]DocumentHandlerInterface
	mammothPath string // Mammoth脚本路径
	debug       bool   // 是否启用调试
	tracer      *Tracer
}

// DocumentHandlerInterface 文档处理器接口
type DocumentHandlerInterface interface {
	Name() string
	Extensions() []string
	CanHandle(ext string) bool
	ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error)
	GetMetadata(filePath string) (*Metadata, error)
}

// NewDocumentHandler 创建文档处理器
func NewDocumentHandler() *DocumentHandler {
	return NewDocumentHandlerWithOptions(false)
}

// NewDocumentHandlerWithOptions 创建文档处理器（带选项）
func NewDocumentHandlerWithOptions(debug bool) *DocumentHandler {
	h := &DocumentHandler{
		handlers:    make(map[string]DocumentHandlerInterface),
		mammothPath: "scripts/mammoth/docx2html.js",
		debug:       debug,
		tracer:      NewTracer(),
	}

	// 启用追踪
	if debug {
		h.tracer.Enable()
		Debug("handler", "DocumentHandler created with debug mode", map[string]interface{}{
			"debug": debug,
		})
	}

	// 注册处理器
	h.registerHandler(newDocxHandler())
	h.registerHandler(newExcelHandler())
	h.registerHandler(newPptxHandler())
	h.registerHandler(newPdfHandler())
	h.registerHandler(newImageHandler())
	h.registerHandler(newTextHandler())
	h.registerHandler(newCsvHandler())

	return h
}

// registerHandler 注册处理器
func (h *DocumentHandler) registerHandler(handler DocumentHandlerInterface) {
	h.handlers[handler.Name()] = handler
}

// SetMammothPath 设置Mammoth脚本路径
func (h *DocumentHandler) SetMammothPath(path string) {
	h.mammothPath = path
}

// Convert 转换文档为HTML
func (h *DocumentHandler) Convert(filePath string, opts ...*ConversionOptions) (*HTMLResult, error) {
	// 开始追踪
	span := h.tracer.StartSpan("Convert", WithTag("file", filePath))
	defer func() {
		h.tracer.StopSpan(span, nil)
	}()

	start := time.Now()

	ext := strings.ToLower(filepath.Ext(filePath))
	ext = strings.TrimPrefix(ext, ".")

	if h.debug {
		Debug("convert", "Starting conversion", map[string]interface{}{
			"file":   filePath,
			"format": ext,
		})
	}

	// 查找处理器
	for _, handler := range h.handlers {
		if handler.CanHandle(ext) {
			options := DefaultOptions
			if len(opts) > 0 && opts[0] != nil {
				options = opts[0]
			}

			// 追踪处理器执行
			handlerSpan := h.tracer.StartSpan(handler.Name(), WithTag("handler", handler.Name()))
			result, err := handler.ToHTML(filePath, options)
			h.tracer.StopSpan(handlerSpan, err)

			// 记录性能
			duration := time.Since(start)
			GetLogger().RecordPerf("Convert:"+ext, duration)

			if h.debug {
				Debug("convert", "Conversion completed", map[string]interface{}{
					"file":     filePath,
					"format":   ext,
					"handler":  handler.Name(),
					"duration": duration.String(),
					"success":  err == nil,
				})
			}

			if err != nil {
				Error("convert", "Conversion failed", map[string]interface{}{
					"file":    filePath,
					"handler": handler.Name(),
					"error":   err.Error(),
				})
			}

			return result, err
		}
	}

	if h.debug {
		Warn("convert", "Unsupported file format", map[string]interface{}{
			"file":   filePath,
			"format": ext,
		})
	}

	return nil, NewError(ErrUnsupportedType, fmt.Sprintf("不支持的文件格式: .%s", ext), nil)
}

// ConvertURL 从URL下载并转换
func (h *DocumentHandler) ConvertURL(url string, opts ...*ConversionOptions) (*HTMLResult, error) {
	// 开始追踪
	span := h.tracer.StartSpan("ConvertURL", WithTag("url", url))
	defer func() {
		h.tracer.StopSpan(span, nil)
	}()

	start := time.Now()

	if h.debug {
		Debug("convert", "Downloading from URL", map[string]interface{}{
			"url": url,
		})
	}

	// 下载文件
	resp, err := http.Get(url)
	if err != nil {
		Error("download", "Failed to download file", map[string]interface{}{
			"url":   url,
			"error": err.Error(),
		})
		return nil, NewError(ErrConversionFailed, "下载文件失败: "+err.Error(), err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		Error("download", "HTTP error", map[string]interface{}{
			"url":    url,
			"status": resp.StatusCode,
		})
		return nil, NewError(ErrConversionFailed, fmt.Sprintf("HTTP错误: %d", resp.StatusCode), nil)
	}

	// 创建临时文件
	tmpFile, err := os.CreateTemp("", "office-*"+filepath.Ext(url))
	if err != nil {
		Error("file", "Failed to create temp file", map[string]interface{}{
			"error": err.Error(),
		})
		return nil, NewError(ErrConversionFailed, "创建临时文件失败: "+err.Error(), err)
	}
	defer os.Remove(tmpFile.Name())
	defer tmpFile.Close()

	// 复制内容
	if _, err := io.Copy(tmpFile, resp.Body); err != nil {
		Error("file", "Failed to save file", map[string]interface{}{
			"error": err.Error(),
		})
		return nil, NewError(ErrConversionFailed, "保存文件失败: "+err.Error(), err)
	}

	duration := time.Since(start)
	GetLogger().RecordPerf("ConvertURL", duration)

	if h.debug {
		Debug("convert", "Download completed, starting conversion", map[string]interface{}{
			"url":      url,
			"duration": duration.String(),
		})
	}

	return h.Convert(tmpFile.Name(), opts...)
}

// GetMetadata 获取文档元信息
func (h *DocumentHandler) GetMetadata(filePath string) (*Metadata, error) {
	ext := strings.ToLower(filepath.Ext(filePath))
	ext = strings.TrimPrefix(ext, ".")

	if h.debug {
		Debug("metadata", "Getting metadata", map[string]interface{}{
			"file":   filePath,
			"format": ext,
		})
	}

	for _, handler := range h.handlers {
		if handler.CanHandle(ext) {
			return handler.GetMetadata(filePath)
		}
	}

	return nil, NewError(ErrUnsupportedType, fmt.Sprintf("不支持的文件格式: .%s", ext), nil)
}

// EnableDebug 启用调试模式
func (h *DocumentHandler) EnableDebug() {
	h.debug = true
	h.tracer.Enable()
	Debug("handler", "Debug mode enabled", nil)
}

// DisableDebug 禁用调试模式
func (h *DocumentHandler) DisableDebug() {
	h.debug = false
	h.tracer.Disable()
	Info("handler", "Debug mode disabled", nil)
}

// IsDebugEnabled 检查调试模式是否启用
func (h *DocumentHandler) IsDebugEnabled() bool {
	return h.debug
}

// GetDebugInfo 获取调试信息
func (h *DocumentHandler) GetDebugInfo() DebugInfo {
	return DebugInfo{
		DebugEnabled: h.debug,
		TracerStats:  h.tracer.GetStats(),
		PerfStats:    GetLogger().GetAllPerfStats(),
		ActiveSpans:  h.tracer.GetActiveSpans(),
		RecentLogs:   GetLogger().GetEntries(DebugLevelDebug, "", 50),
	}
}

// DebugInfo 调试信息
type DebugInfo struct {
	DebugEnabled bool         `json:"debug_enabled"`
	TracerStats  TracerStats `json:"tracer_stats"`
	PerfStats    []*PerfStats `json:"perf_stats"`
	ActiveSpans  []*Span     `json:"active_spans"`
	RecentLogs   []LogEntry  `json:"recent_logs"`
}

// GetSupportedFormats 获取支持的格式
func (h *DocumentHandler) GetSupportedFormats() []string {
	var formats []string
	for _, handler := range h.handlers {
		formats = append(formats, handler.Extensions()...)
	}
	return formats
}

// Name 返回处理器名称
func (h *DocumentHandler) Name() string {
	return "main"
}

// Extensions 返回支持的扩展名
func (h *DocumentHandler) Extensions() []string {
	return h.GetSupportedFormats()
}

// CanHandle 检查是否支持该格式
func (h *DocumentHandler) CanHandle(ext string) bool {
	for _, handler := range h.handlers {
		if handler.CanHandle(ext) {
			return true
		}
	}
	return false
}

// ToHTML 转换为HTML (由具体处理器实现)
func (h *DocumentHandler) ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	return h.Convert(filePath, opts)
}

// 并发安全的文档处理器池
type handlerPool struct {
	pool sync.Pool
}

func newHandlerPool() *handlerPool {
	return &handlerPool{
		pool: sync.Pool{
			New: func() interface{} {
				return NewDocumentHandler()
			},
		},
	}
}

var defaultPool = newHandlerPool()

// GetFromPool 从池获取处理器
func GetFromPool() *DocumentHandler {
	return defaultPool.pool.Get().(*DocumentHandler)
}

// PutToPool 归还处理器到池
func PutToPool(h *DocumentHandler) {
	defaultPool.pool.Put(h)
}
