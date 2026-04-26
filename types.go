package office

import (
	"strings"
	"time"
)

// DocumentType 文档类型
type DocumentType string

const (
	TypeDOCX  DocumentType = "docx"  // Word文档
	TypeXLSX  DocumentType = "xlsx"  // Excel表格
	TypeXLS   DocumentType = "xls"  // Excel 97-2003
	TypePPTX  DocumentType = "pptx"  // PowerPoint演示
	TypePDF   DocumentType = "pdf"   // PDF文档
	TypeImage DocumentType = "image" // 图片
	TypeText  DocumentType = "text"  // 文本
	TypeCSV   DocumentType = "csv"   // CSV表格
)

// FormatInfo 格式信息
type FormatInfo struct {
	Extension  string       // 扩展名
	Name       string       // 格式名称
	Category   string       // 分类
	Type       DocumentType // 类型
	Processor  string      // 处理器
}

// AllFormats 所有支持的格式
var AllFormats = []FormatInfo{
	// Excel表格
	{"xlsx", "Excel 2007+", "表格", TypeXLSX, "excelize"},
	{"xls", "Excel 97-2003", "表格", TypeXLS, "excelize"},
	{"xlsm", "Excel 宏文件", "表格", TypeXLSX, "excelize"},
	{"et", "WPS表格", "表格", TypeXLSX, "excelize"},
	{"ett", "WPS表格模板", "表格", TypeXLSX, "excelize"},

	// Word文档
	{"docx", "Word 2007+", "文档", TypeDOCX, "mammoth"},
	{"doc", "Word 97-2003", "文档", TypeDOCX, "mammoth"},
	{"wps", "WPS文字", "文档", TypeDOCX, "mammoth"},
	{"docm", "Word 宏文件", "文档", TypeDOCX, "mammoth"},
	{"dotx", "Word模板", "文档", TypeDOCX, "mammoth"},
	{"dot", "Word 97模板", "文档", TypeDOCX, "mammoth"},

	// PowerPoint演示
	{"pptx", "PowerPoint 2007+", "演示", TypePPTX, "mammoth"},
	{"ppt", "PowerPoint 97-2003", "演示", TypePPTX, "mammoth"},
	{"dps", "WPS演示", "演示", TypePPTX, "mammoth"},
	{"pptm", "PowerPoint宏文件", "演示", TypePPTX, "mammoth"},

	// PDF
	{"pdf", "PDF文档", "文档", TypePDF, "pdf.js"},

	// CSV
	{"csv", "CSV文件", "表格", TypeCSV, "内置"},
	{"tsv", "TSV文件", "表格", TypeCSV, "内置"},

	// 图片
	{"jpg", "JPEG图片", "图片", TypeImage, "内置"},
	{"jpeg", "JPEG图片", "图片", TypeImage, "内置"},
	{"png", "PNG图片", "图片", TypeImage, "内置"},
	{"gif", "GIF动画", "图片", TypeImage, "内置"},
	{"bmp", "位图", "图片", TypeImage, "内置"},
	{"webp", "WebP图片", "图片", TypeImage, "内置"},
	{"svg", "矢量图", "图片", TypeImage, "内置"},
	{"ico", "图标", "图片", TypeImage, "内置"},

	// 文本
	{"txt", "纯文本", "文本", TypeText, "内置"},
	{"text", "纯文本", "文本", TypeText, "内置"},
	{"md", "Markdown", "文本", TypeText, "内置"},
	{"markdown", "Markdown", "文本", TypeText, "内置"},
	{"log", "日志文件", "文本", TypeText, "内置"},
	{"json", "JSON", "文本", TypeText, "内置"},
	{"xml", "XML", "文本", TypeText, "内置"},
	{"html", "HTML", "文本", TypeText, "内置"},
	{"css", "CSS", "文本", TypeText, "内置"},
	{"js", "JavaScript", "文本", TypeText, "内置"},
	{"go", "Go", "文本", TypeText, "内置"},
	{"java", "Java", "文本", TypeText, "内置"},
	{"py", "Python", "文本", TypeText, "内置"},
	{"sql", "SQL", "文本", TypeText, "内置"},
	{"yaml", "YAML", "文本", TypeText, "内置"},
	{"yml", "YAML", "文本", TypeText, "内置"},
	{"toml", "TOML", "文本", TypeText, "内置"},
	{"ini", "INI配置", "文本", TypeText, "内置"},
	{"conf", "配置文件", "文本", TypeText, "内置"},
}

// GetFormatInfo 获取格式信息
func GetFormatInfo(ext string) *FormatInfo {
	ext = strings.ToLower(strings.TrimPrefix(ext, "."))
	for _, f := range AllFormats {
		if f.Extension == ext {
			return &f
		}
	}
	return nil
}

// HTMLResult HTML转换结果
type HTMLResult struct {
	HTML       string            // HTML内容
	CSS        string            // 额外CSS
	JS         string            // 额外JS
	Images     map[string][]byte  // 内嵌图片: name -> data
	Metadata   *Metadata         // 文档元信息
	PageCount  int               // 页数
	WordCount  int               // 字数
	Format     DocumentType      // 文档类型
}

// Metadata 文档元信息
type Metadata struct {
	Title       string    // 标题
	Author      string    // 作者
	Subject     string    // 主题
	Keywords    string    // 关键词
	Created     time.Time // 创建时间
	Modified    time.Time // 修改时间
	FileSize    int64     // 文件大小
	PageCount   int       // 页数
	SheetCount  int       // 工作表数量 (Excel)
	SlideCount  int       // 幻灯片数量 (PPT)
	WordCount   int       // 字数
}

// ConversionOptions 转换选项
type ConversionOptions struct {
	EmbedImages   bool // 内嵌图片到HTML
	MaxWidth      int  // 最大宽度 (px)
	Theme         string // 主题: "light" / "dark"
	ShowPageBreaks bool // 显示分页符
	FontScale     float64 // 字体缩放比例
}

// DefaultOptions 默认选项
var DefaultOptions = &ConversionOptions{
	EmbedImages:   true,
	MaxWidth:      1200,
	Theme:         "light",
	ShowPageBreaks: false,
	FontScale:     1.0,
}

// ErrorCode 错误码
type ErrorCode string

const (
	ErrFileNotFound     ErrorCode = "FILE_NOT_FOUND"
	ErrUnsupportedType ErrorCode = "UNSUPPORTED_TYPE"
	ErrConversionFailed ErrorCode = "CONVERSION_FAILED"
	ErrTimeout         ErrorCode = "TIMEOUT"
	ErrInvalidFormat   ErrorCode = "INVALID_FORMAT"
)

// OfficeError 错误结构
type OfficeError struct {
	Code    ErrorCode
	Message string
	Err     error
}

func (e *OfficeError) Error() string {
	return e.Message
}

func (e *OfficeError) Unwrap() error {
	return e.Err
}

// NewError 创建错误
func NewError(code ErrorCode, msg string, err error) *OfficeError {
	return &OfficeError{
		Code:    code,
		Message: msg,
		Err:     err,
	}
}
