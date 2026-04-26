package office

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

// pdfHandler PDF文档处理器
type pdfHandler struct {
	scriptPath string
}

// newPdfHandler 创建PDF处理器
func newPdfHandler() *pdfHandler {
	return &pdfHandler{
		scriptPath: "scripts/pdfjs/pdf2html.js",
	}
}

// Name 返回处理器名称
func (h *pdfHandler) Name() string {
	return "pdf"
}

// Extensions 返回支持的扩展名
func (h *pdfHandler) Extensions() []string {
	return []string{"pdf"}
}

// CanHandle 检查是否支持该格式
func (h *pdfHandler) CanHandle(ext string) bool {
	ext = strings.ToLower(ext)
	for _, e := range h.Extensions() {
		if ext == e {
			return true
		}
	}
	return false
}

// ToHTML 转换为HTML
func (h *pdfHandler) ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, NewError(ErrFileNotFound, "文件不存在: "+filePath, err)
	}

	// 优先使用本地 PDF.js 渲染
	if html, err := h.convertWithLocalPDFJS(filePath, opts); err == nil {
		return &HTMLResult{
			HTML: html,
			Metadata: &Metadata{
				Title:    filepath.Base(filePath),
				FileSize: getFileSize(filePath),
				Modified: getFileModTime(filePath),
			},
			Format: TypePDF,
		}, nil
	}

	// 降级到 CDN 方式（需要网络连接）
	html, err := h.convertWithCDN(filePath, opts)
	if err != nil {
		return nil, err
	}
	return &HTMLResult{
		HTML: html,
		Metadata: &Metadata{
			Title:    filepath.Base(filePath),
			FileSize: getFileSize(filePath),
			Modified: getFileModTime(filePath),
		},
		Format: TypePDF,
	}, nil
}

// convertWithLocalPDFJS 使用本地 PDF.js 转换
func (h *pdfHandler) convertWithLocalPDFJS(filePath string, opts *ConversionOptions) (string, error) {
	// 检查 Node.js 是否可用
	if !isNodeAvailable() {
		return "", fmt.Errorf("Node.js not available")
	}

	// 检查脚本是否存在
	scriptPath := h.scriptPath
	if _, err := os.Stat(scriptPath); os.IsNotExist(err) {
		return "", fmt.Errorf("PDF.js script not found")
	}

	// 检查 node_modules 是否存在
	modulePath := filepath.Join(filepath.Dir(scriptPath), "node_modules", "pdfjs-dist")
	if _, err := os.Stat(modulePath); os.IsNotExist(err) {
		return "", fmt.Errorf("pdfjs-dist not installed, run: cd scripts/pdfjs && npm install")
	}

	// 执行转换
	cmd := exec.Command("node", scriptPath, filePath)
	output, err := cmd.Output()
	if err != nil {
		return "", err
	}

	return string(output), nil
}

// convertWithCDN 使用 CDN 方式转换（需要网络连接）
func (h *pdfHandler) convertWithCDN(filePath string, opts *ConversionOptions) (string, error) {
	pdfData, err := os.ReadFile(filePath)
	if err != nil {
		return "", NewError(ErrConversionFailed, "读取PDF失败: "+err.Error(), err)
	}

	css := h.getCSS(opts)
	encoded := base64.StdEncoding.EncodeToString(pdfData)

	return fmt.Sprintf(`<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf_viewer.min.css">
<style>%s</style>
</head>
<body>
<div class="pdf-container">
  <div class="pdf-header">
    <span class="pdf-title">%s</span>
  </div>
  <div id="pdf-viewer"></div>
  <div class="pdf-fallback">
    <p>提示: 如果PDF无法预览，请下载查看 <a href="file://%s" download>下载PDF</a></p>
  </div>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script>
const pdfData = atob('%s');
const pdfArray = new Uint8Array(pdfData.length);
for(let i = 0; i < pdfData.length; i++) pdfArray[i] = pdfData.charCodeAt(i);
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
try {
    const pdf = await pdfjsLib.getDocument({data: pdfArray}).promise;
    const viewer = document.getElementById('pdf-viewer');
    for(let page = 1; page <= pdf.numPages; page++) {
        const canvas = document.createElement('canvas');
        canvas.className = 'pdf-page';
        const ctx = canvas.getContext('2d');
        const scale = 1.5;
        const viewport = (await pdf.getPage(page)).getViewport({scale});
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        viewer.appendChild(canvas);
        await (await pdf.getPage(page)).render({canvasContext: ctx, viewport}).promise;
    }
} catch(e) {
    viewer.innerHTML = '<p class="error">PDF预览失败: '+e.message+'</p>';
}
</script>
</body>
</html>`, filepath.Base(filePath), css, filepath.Base(filePath), filePath, encoded), nil
}

// getCSS 获取CSS样式
func (h *pdfHandler) getCSS(opts *ConversionOptions) string {
	return `
* { box-sizing: border-box; margin: 0; padding: 0; }
body { background: #525252; font-family: Arial, sans-serif; }
.pdf-container { max-width: 900px; margin: 0 auto; }
.pdf-header { background: #323232; color: white; padding: 12px 20px; position: sticky; top: 0; z-index: 100; }
.pdf-title { font-size: 14px; }
.pdf-page { display: block; margin: 20px auto; box-shadow: 0 4px 12px rgba(0,0,0,0.3); background: white; }
.pdf-fallback { background: #323232; color: white; padding: 12px 20px; text-align: center; }
.pdf-fallback a { color: #4da6ff; }
.error { color: #ff6b6b; padding: 20px; text-align: center; background: white; margin: 20px; }
`
}

// GetMetadata 获取文档元信息
func (h *pdfHandler) GetMetadata(filePath string) (*Metadata, error) {
	return &Metadata{
		Title:    filepath.Base(filePath),
		FileSize: getFileSize(filePath),
		Modified: getFileModTime(filePath),
	}, nil
}

// isNodeAvailable 检查 Node.js 是否可用
func isNodeAvailable() bool {
	cmd := exec.Command("node", "--version")
	return cmd.Run() == nil
}

// PDFMetadata PDF元信息结构
type PDFMetadata struct {
	Title    string `json:"title"`
	Author   string `json:"author"`
	Subject  string `json:"subject"`
	Creator  string `json:"creator"`
}

// ExtractMetadata 提取PDF元信息
func (h *pdfHandler) ExtractMetadata(filePath string) (*PDFMetadata, error) {
	// 使用 pdfinfo 或类似工具（如果可用）
	// 这里先返回基本元信息
	meta, err := h.GetMetadata(filePath)
	if err != nil {
		return nil, err
	}
	return &PDFMetadata{
		Title: meta.Title,
	}, nil
}

// MarshalJSON 实现 json.Marshaler
func (m *PDFMetadata) MarshalJSON() ([]byte, error) {
	return json.Marshal(map[string]string{
		"title":   m.Title,
		"author":  m.Author,
		"subject": m.Subject,
		"creator": m.Creator,
	})
}
