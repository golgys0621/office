package office

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"strings"
)

// docxHandler DOCX文档处理器
type docxHandler struct {
	mammothPath string
}

// newDocxHandler 创建DOCX处理器
func newDocxHandler() *docxHandler {
	return &docxHandler{
		mammothPath: "scripts/mammoth/docx2html.js",
	}
}

// Name 返回处理器名称
func (h *docxHandler) Name() string {
	return "docx"
}

// Extensions 返回支持的扩展名
func (h *docxHandler) Extensions() []string {
	return []string{"docx", "wps", "docm", "dotx", "dot"}
}

// CanHandle 检查是否支持该格式
func (h *docxHandler) CanHandle(ext string) bool {
	ext = strings.ToLower(ext)
	for _, e := range h.Extensions() {
		if ext == e {
			return true
		}
	}
	return false
}

// ToHTML 转换为HTML
func (h *docxHandler) ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	// 检查文件是否存在
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, NewError(ErrFileNotFound, "文件不存在: "+filePath, err)
	}

	// 尝试使用Mammoth转换
	html, err := h.convertWithMammoth(filePath)
	if err != nil {
		// 如果Mammoth失败，回退到内置解析
		return h.fallbackToBuiltin(filePath, opts)
	}

	// 提取元信息
	metadata, _ := h.GetMetadata(filePath)

	// 包装HTML
	wrappedHTML := h.wrapHTML(html, filepath.Base(filePath), opts)

	return &HTMLResult{
		HTML:     wrappedHTML,
		Metadata: metadata,
		Format:   TypeDOCX,
	}, nil
}

// convertWithMammoth 使用Mammoth转换
func (h *docxHandler) convertWithMammoth(filePath string) (string, error) {
	// 查找Node.js
	nodePath, err := exec.LookPath("node")
	if err != nil {
		return "", fmt.Errorf("Node.js未安装: %v", err)
	}

	// 查找Mammoth脚本
	scriptPath := h.mammothPath
	if !filepath.IsAbs(scriptPath) {
		// 相对于当前工作目录
		cwd, _ := os.Getwd()
		scriptPath = filepath.Join(cwd, scriptPath)
	}

	if _, err := os.Stat(scriptPath); os.IsNotExist(err) {
		return "", fmt.Errorf("Mammoth脚本不存在: %s", scriptPath)
	}

	// 创建临时输出文件
	tmpOut, err := os.CreateTemp("", "mammoth-out-*.json")
	if err != nil {
		return "", fmt.Errorf("创建临时文件失败: %v", err)
	}
	tmpOut.Close()
	defer os.Remove(tmpOut.Name())

	// 执行Mammoth
	cmd := exec.Command(nodePath, scriptPath, filePath, tmpOut.Name())
	output, err := cmd.CombinedOutput()

	if err != nil {
		return "", fmt.Errorf("Mammoth执行失败: %v, output: %s", err, string(output))
	}

	// 读取结果
	resultData, err := os.ReadFile(tmpOut.Name())
	if err != nil {
		return "", fmt.Errorf("读取结果失败: %v", err)
	}

	// 解析JSON
	var mammothResult struct {
		Success  bool     `json:"success"`
		HTML     string   `json:"html"`
		Messages []string `json:"messages"`
		Warnings []string `json:"warnings"`
		Stats    struct {
			Paragraphs int `json:"paragraphs"`
			Images    int `json:"images"`
			Tables    int `json:"tables"`
		} `json:"stats"`
		Error string `json:"error"`
	}

	if err := json.Unmarshal(resultData, &mammothResult); err != nil {
		return "", fmt.Errorf("解析结果失败: %v", err)
	}

	if !mammothResult.Success {
		return "", fmt.Errorf("Mammoth转换失败: %s", mammothResult.Error)
	}

	return mammothResult.HTML, nil
}

// fallbackToBuiltin 内置解析（备选方案）
func (h *docxHandler) fallbackToBuiltin(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	// 这里复用之前的正则解析逻辑
	return nil, fmt.Errorf("Mammoth未安装，使用内置解析器")
}

// wrapHTML 包装HTML
func (h *docxHandler) wrapHTML(html, fileName string, opts *ConversionOptions) string {
	css := h.getCSS(opts)
	
	return fmt.Sprintf(`<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s</title>
<style>
%s
</style>
</head>
<body>
<div class="office-container">
%s
</div>
</body>
</html>`, fileName, css, html)
}

// getCSS 获取CSS样式
func (h *docxHandler) getCSS(opts *ConversionOptions) string {
	theme := "light"
	if opts != nil && opts.Theme != "" {
		theme = opts.Theme
	}
	
	var bgColor, textColor string
	if theme == "dark" {
		bgColor = "#1e1e1e"
		textColor = "#d4d4d4"
	} else {
		bgColor = "#f3f3f3"
		textColor = "#333333"
	}

	return fmt.Sprintf(`
* { box-sizing: border-box; margin: 0; padding: 0; }
body { 
    font-family: 'Microsoft YaHei', '微软雅黑', 'Segoe UI', Arial, sans-serif; 
    background: %s; 
    color: %s;
    line-height: 1.6;
    padding: 20px;
}
.office-container { 
    max-width: %dpx; 
    margin: 0 auto; 
    background: white; 
    padding: 40px 60px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    border-radius: 4px;
}
.office-container h1 { font-size: 24pt; margin: 0 0 20px 0; text-align: center; }
.office-container h2 { font-size: 18pt; margin: 20px 0 12px 0; border-bottom: 1px solid #ddd; padding-bottom: 8px; }
.office-container h3 { font-size: 14pt; margin: 16px 0 8px 0; }
.office-container p { margin: 8px 0; text-indent: 2em; }
.office-container ul, .office-container ol { margin: 10px 0 10px 30px; }
.office-container li { margin: 5px 0; }
.office-container table { border-collapse: collapse; width: 100%%; margin: 16px 0; }
.office-container td, .office-container th { border: 1px solid #ddd; padding: 8px 12px; }
.office-container th { background: #f5f5f5; font-weight: bold; }
.office-container img { max-width: 100%%; height: auto; margin: 10px 0; }
.office-container blockquote { 
    border-left: 4px solid #007bff; 
    padding: 10px 15px; 
    margin: 16px 0; 
    background: #f8f9fa;
}
`, bgColor, textColor, opts.MaxWidth)
}

// GetMetadata 获取文档元信息
func (h *docxHandler) GetMetadata(filePath string) (*Metadata, error) {
	fileInfo, err := os.Stat(filePath)
	if err != nil {
		return nil, NewError(ErrFileNotFound, "文件不存在: "+filePath, err)
	}

	// 尝试使用Mammoth获取统计信息
	stats, err := h.getMammothStats(filePath)
	if err != nil {
		// 回退到基础统计
		stats = &docxStats{}
	}

	_ = stats // 保留以备后用

	return &Metadata{
		Title:    filepath.Base(filePath),
		FileSize: fileInfo.Size(),
		Created:  fileInfo.ModTime(),
		Modified: fileInfo.ModTime(),
	}, nil
}

// docxStats DOCX统计信息
type docxStats struct {
	Characters  int
	Words       int
	Paragraphs  int
	Images      int
	Tables      int
}

// getMammothStats 使用Mammoth获取统计信息
func (h *docxHandler) getMammothStats(filePath string) (*docxStats, error) {
	nodePath, err := exec.LookPath("node")
	if err != nil {
		return nil, err
	}

	scriptPath := h.mammothPath
	if !filepath.IsAbs(scriptPath) {
		cwd, _ := os.Getwd()
		scriptPath = filepath.Join(cwd, scriptPath)
	}

	tmpOut, err := os.CreateTemp("", "mammoth-meta-*.json")
	if err != nil {
		return nil, err
	}
	tmpOut.Close()
	defer os.Remove(tmpOut.Name())

	cmd := exec.Command(nodePath, scriptPath, "--metadata", filePath)
	output, err := cmd.CombinedOutput()
	if err != nil {
		return nil, err
	}

	var result struct {
		Success bool `json:"success"`
		Stats   struct {
			Characters  int `json:"characters"`
			Words       int `json:"words"`
			Paragraphs  int `json:"paragraphs"`
		} `json:"stats"`
	}

	if err := json.Unmarshal(output, &result); err != nil {
		return nil, err
	}

	return &docxStats{
		Characters: result.Stats.Characters,
		Words:      result.Stats.Words,
		Paragraphs: result.Stats.Paragraphs,
	}, nil
}

// extractDocxText 提取DOCX纯文本（用于统计）
func extractDocxText(filePath string) (string, error) {
	return extractZipText(filePath, "word/document.xml")
}

// extractZipText 从ZIP中提取特定文件的文本
func extractZipText(filePath, targetFile string) (string, error) {
	// 简单实现，实际应该用zip.Reader
	return "", nil
}

// countWords 统计字数
func countWords(text string) int {
	// 中文字符每个算1字
	// 英文单词按空格分割
	chineseChars := regexp.MustCompile(`[\u4e00-\u9fa5]`).FindAllString(text, -1)
	englishWords := regexp.MustCompile(`[a-zA-Z]+`).FindAllString(text, -1)
	
	return len(chineseChars) + len(englishWords)
}

// Ensure HTMLResult has WordCount
func init() {
	// This will be called at package init
}

// GetWordCount 获取文档字数
func GetWordCount(filePath string) (int, error) {
	text, err := extractDocxText(filePath)
	if err != nil {
		return 0, err
	}
	return countWords(text), nil
}
