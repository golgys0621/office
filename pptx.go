package office

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

// pptxHandler PPTX文档处理器
type pptxHandler struct {
	scriptPath string
}

// newPptxHandler 创建PPTX处理器
func newPptxHandler() *pptxHandler {
	return &pptxHandler{
		scriptPath: "scripts/mammoth/pptx2html.js",
	}
}

// Name 返回处理器名称
func (h *pptxHandler) Name() string {
	return "pptx"
}

// Extensions 返回支持的扩展名
func (h *pptxHandler) Extensions() []string {
	return []string{"pptx", "dps", "ppt", "pptm"}
}

// CanHandle 检查是否支持该格式
func (h *pptxHandler) CanHandle(ext string) bool {
	ext = strings.ToLower(ext)
	for _, e := range h.Extensions() {
		if ext == e {
			return true
		}
	}
	return false
}

// ToHTML 转换为HTML
func (h *pptxHandler) ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, NewError(ErrFileNotFound, "文件不存在: "+filePath, err)
	}

	// 尝试使用Mammoth转换
	html, err := h.convertWithMammoth(filePath)
	if err != nil {
		// 回退到内置解析
		return h.fallbackToBuiltin(filePath, opts)
	}

	// 包装HTML
	wrappedHTML := h.wrapHTML(html, filepath.Base(filePath), opts)

	return &HTMLResult{
		HTML: wrappedHTML,
		Metadata: &Metadata{
			Title:      filepath.Base(filePath),
			FileSize:   getFileSize(filePath),
			Modified:   getFileModTime(filePath),
		},
		Format: TypePPTX,
	}, nil
}

// convertWithMammoth 使用Mammoth转换
func (h *pptxHandler) convertWithMammoth(filePath string) (string, error) {
	nodePath, err := exec.LookPath("node")
	if err != nil {
		return "", fmt.Errorf("Node.js未安装")
	}

	scriptPath := h.scriptPath
	if !filepath.IsAbs(scriptPath) {
		cwd, _ := os.Getwd()
		scriptPath = filepath.Join(cwd, scriptPath)
	}

	tmpOut, err := os.CreateTemp("", "pptx-out-*.json")
	if err != nil {
		return "", err
	}
	tmpOut.Close()
	defer os.Remove(tmpOut.Name())

	cmd := exec.Command(nodePath, scriptPath, filePath)
	output, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("Mammoth执行失败: %v, output: %s", err, string(output))
	}

	var result struct {
		Success bool   `json:"success"`
		HTML    string `json:"html"`
		Stats   struct {
			Slides int `json:"slides"`
		} `json:"stats"`
		Error string `json:"error"`
	}

	if err := json.Unmarshal(output, &result); err != nil {
		return "", fmt.Errorf("解析结果失败: %v", err)
	}

	if !result.Success {
		return "", fmt.Errorf("PPTX转换失败: %s", result.Error)
	}

	return result.HTML, nil
}

// fallbackToBuiltin 内置解析（备选方案）
func (h *pptxHandler) fallbackToBuiltin(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	// 使用ZIP解析PPTX基础结构
	return h.parseBasicStructure(filePath, opts)
}

// parseBasicStructure 解析PPTX基本结构
func (h *pptxHandler) parseBasicStructure(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	// 简化实现，实际应该用zip解析
	return nil, fmt.Errorf("Mammoth未安装")
}

// wrapHTML 包装HTML
func (h *pptxHandler) wrapHTML(html, fileName string, opts *ConversionOptions) string {
	css := h.getCSS(opts)
	js := h.getJS()

	return fmt.Sprintf(`<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s</title>
<style>%s</style>
</head>
<body>
<div class="pptx-container">
  <div class="pptx-header">%s</div>
  <div class="pptx-slides">%s</div>
</div>
<script>%s</script>
</body>
</html>`, fileName, css, fileName, html, js)
}

// getCSS 获取CSS样式
func (h *pptxHandler) getCSS(opts *ConversionOptions) string {
	maxWidth := 1200
	if opts != nil && opts.MaxWidth > 0 {
		maxWidth = opts.MaxWidth
	}

	return fmt.Sprintf(`
* { box-sizing: border-box; margin: 0; padding: 0; }
body { 
    font-family: 'Microsoft YaHei', Arial, sans-serif; 
    background: #1e1e1e; 
    color: white;
}
.pptx-container { 
    max-width: %dpx; 
    margin: 0 auto; 
}
.pptx-header { 
    background: #d24726; 
    color: white; 
    padding: 12px 20px; 
    font-size: 16px;
    font-weight: bold;
    position: sticky;
    top: 0;
    z-index: 100;
}
.pptx-slides { 
    background: #1e1e1e; 
    padding: 20px;
}
.slide { 
    background: white; 
    color: #333; 
    margin-bottom: 20px; 
    border-radius: 8px; 
    overflow: hidden;
    box-shadow: 0 4px 12px rgba(0,0,0,0.3);
}
.slide-content { 
    padding: 40px; 
    min-height: 400px;
}
.slide-title { 
    font-size: 28pt; 
    font-weight: bold; 
    margin-bottom: 20px;
    color: #333;
}
.slide-text { 
    font-size: 14pt; 
    line-height: 1.8;
    color: #333;
}
.slide-nav {
    display: flex;
    justify-content: center;
    gap: 10px;
    padding: 20px;
    background: #2d2d2d;
    position: fixed;
    bottom: 0;
    width: 100%%;
}
.nav-btn {
    padding: 10px 30px;
    background: #d24726;
    color: white;
    border: none;
    border-radius: 4px;
    cursor: pointer;
}
.nav-btn:hover { background: #e55a38; }
`, maxWidth)
}

// getJS 获取JavaScript代码
func (h *pptxHandler) getJS() string {
	return `
let currentSlide = 0;
function goToSlide(index) {
    const slides = document.querySelectorAll('.slide');
    if (index >= 0 && index < slides.length) {
        currentSlide = index;
        slides[index].scrollIntoView({ behavior: 'smooth' });
    }
}
function nextSlide() { goToSlide(currentSlide + 1); }
function prevSlide() { goToSlide(currentSlide - 1); }
document.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowRight' || e.key === ' ') nextSlide();
    if (e.key === 'ArrowLeft') prevSlide();
});
`
}

// GetMetadata 获取文档元信息
func (h *pptxHandler) GetMetadata(filePath string) (*Metadata, error) {
	return &Metadata{
		Title:    filepath.Base(filePath),
		FileSize: getFileSize(filePath),
		Modified: getFileModTime(filePath),
	}, nil
}
