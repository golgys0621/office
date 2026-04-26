package office

import (
	"net/http"
	"os"
)

// PreviewIntegration 与现有预览系统的集成
type PreviewIntegration struct {
	*Integration
	tempDir string
}

// NewPreviewIntegration 创建预览集成
func NewPreviewIntegration() *PreviewIntegration {
	tempDir, _ := os.MkdirTemp("", "office-preview-")
	return &PreviewIntegration{
		Integration: NewIntegration(),
		tempDir:     tempDir,
	}
}

// Cleanup 清理临时文件
func (p *PreviewIntegration) Cleanup() {
	os.RemoveAll(p.tempDir)
}

// ConvertForPreview 转换为预览格式
func (p *PreviewIntegration) ConvertForPreview(filePath string) (string, error) {
	// 设置合适的预览选项
	opts := &ConversionOptions{
		EmbedImages:    true,
		MaxWidth:       1000,
		Theme:          "light",
		ShowPageBreaks: false,
		FontScale:      1.0,
	}

	result, err := p.handler.Convert(filePath, opts)
	if err != nil {
		return "", err
	}
	return result.HTML, nil
}

// ConvertURLForPreview 从URL转换预览
func (p *PreviewIntegration) ConvertURLForPreview(url string) (string, error) {
	opts := &ConversionOptions{
		EmbedImages:    true,
		MaxWidth:       1000,
		Theme:          "light",
		ShowPageBreaks: false,
		FontScale:      1.0,
	}

	result, err := p.handler.ConvertURL(url, opts)
	if err != nil {
		return "", err
	}
	return result.HTML, nil
}

// HTTPHandler 返回一个标准的 http.HandlerFunc
// 使用示例:
//
//	import "github.com/golgys0621/office"
//	http.HandleFunc("/preview", office.PreviewIntegration{}.HTTPHandler())
func (p *PreviewIntegration) HTTPHandler() http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		fileURL := r.URL.Query().Get("url")
		if fileURL == "" {
			http.Error(w, "缺少 url 参数", http.StatusBadRequest)
			return
		}

		html, err := p.ConvertURLForPreview(fileURL)
		if err != nil {
			http.Error(w, "转换失败: "+err.Error(), http.StatusInternalServerError)
			return
		}

		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		w.Write([]byte(html))
	}
}

// PreviewResult 预览结果
type PreviewResult struct {
	HTML      string
	Format    string
	Title     string
	PageCount int
	WordCount int
	FileSize  int64
}

// FormatResult 格式化结果
func FormatResult(result *HTMLResult) *PreviewResult {
	pr := &PreviewResult{
		HTML:   result.HTML,
		Format: string(result.Format),
	}

	if result.Metadata != nil {
		pr.Title = result.Metadata.Title
		pr.FileSize = result.Metadata.FileSize
	}

	return pr
}
