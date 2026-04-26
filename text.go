package office

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// textHandler 文本处理器
type textHandler struct{}

// newTextHandler 创建文本处理器
func newTextHandler() *textHandler {
	return &textHandler{}
}

// Name 返回处理器名称
func (h *textHandler) Name() string {
	return "text"
}

// Extensions 返回支持的扩展名
func (h *textHandler) Extensions() []string {
	return []string{"txt", "text", "log", "md", "markdown", "json", "xml", "html", "css", "js", "go", "java", "py", "c", "cpp", "h", "cs", "rb", "php", "sh", "yaml", "yml", "toml", "ini", "conf", "cfg", "sql"}
}

// CanHandle 检查是否支持该格式
func (h *textHandler) CanHandle(ext string) bool {
	ext = strings.ToLower(ext)
	for _, e := range h.Extensions() {
		if ext == e {
			return true
		}
	}
	return false
}

// ToHTML 转换为HTML
func (h *textHandler) ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, NewError(ErrFileNotFound, "文件不存在: "+filePath, err)
	}

	// 读取文本内容
	content, err := os.ReadFile(filePath)
	if err != nil {
		return nil, NewError(ErrConversionFailed, "读取文件失败: "+err.Error(), err)
	}

	ext := strings.ToLower(filepath.Ext(filePath))
	isMarkdown := ext == ".md" || ext == ".markdown"
	
	var htmlContent string
	if isMarkdown {
		htmlContent = convertMarkdownToHTML(string(content))
	} else {
		htmlContent = escapeAndHighlight(string(content), ext)
	}

	html := h.wrapHTML(htmlContent, filepath.Base(filePath), opts, ext)

	return &HTMLResult{
		HTML: html,
		Metadata: &Metadata{
			Title:    filepath.Base(filePath),
			FileSize: getFileSize(filePath),
			Modified: getFileModTime(filePath),
			WordCount: len(string(content)),
		},
		Format: TypeText,
	}, nil
}

// wrapHTML 包装HTML
func (h *textHandler) wrapHTML(content, fileName string, opts *ConversionOptions, ext string) string {
	css := h.getCSS(ext)
	js := h.getJS(ext)

	return fmt.Sprintf(`<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s</title>
%s
<style>%s</style>
</head>
<body>
<div class="text-container">
  <div class="text-header">
    <span class="text-title">%s</span>
    <span class="text-lang">%s</span>
  </div>
  <pre class="text-content"><code>%s</code></pre>
</div>
%s
</body>
</html>`, fileName, h.getHighlightCSS(ext), css, fileName, strings.TrimPrefix(ext, "."), content, js)
}

// getCSS 获取CSS样式
func (h *textHandler) getCSS(ext string) string {
	return `
* { box-sizing: border-box; margin: 0; padding: 0; }
body { background: #1e1e1e; font-family: 'Consolas', 'Monaco', 'Courier New', monospace; }
.text-container { max-width: 1200px; margin: 0 auto; }
.text-header { 
    background: #333; 
    color: white; 
    padding: 12px 20px; 
    display: flex; 
    justify-content: space-between;
    position: sticky;
    top: 0;
}
.text-title { font-size: 14px; }
.text-lang { font-size: 12px; color: #aaa; text-transform: uppercase; }
.text-content { 
    padding: 20px; 
    overflow-x: auto; 
    line-height: 1.6;
    font-size: 14px;
}
`
}

// getHighlightCSS 获取语法高亮CSS
func (h *textHandler) getHighlightCSS(ext string) string {
	return `<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/vs2015.min.css">`
}

// getJS 获取JavaScript代码
func (h *textHandler) getJS(ext string) string {
	lang := getLanguageClass(ext)
	return fmt.Sprintf(`<script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/highlight.min.js"></script>
<script>hljs.highlightAll();</script>`, lang)
}

// GetMetadata 获取文档元信息
func (h *textHandler) GetMetadata(filePath string) (*Metadata, error) {
	content, _ := os.ReadFile(filePath)
	return &Metadata{
		Title:    filepath.Base(filePath),
		FileSize: getFileSize(filePath),
		Modified: getFileModTime(filePath),
		WordCount: len(string(content)),
	}, nil
}

// escapeAndHighlight HTML转义并标记语言
func escapeAndHighlight(content, ext string) string {
	// HTML转义
	content = strings.ReplaceAll(content, "&", "&amp;")
	content = strings.ReplaceAll(content, "<", "&lt;")
	content = strings.ReplaceAll(content, ">", "&gt;")
	
	return content
}

// convertMarkdownToHTML 简单的Markdown转HTML
func convertMarkdownToHTML(content string) string {
	var result strings.Builder
	lines := strings.Split(content, "\n")
	inCodeBlock := false
	
	for _, line := range lines {
		// 代码块
		if strings.HasPrefix(line, "```") {
			if inCodeBlock {
				result.WriteString("</code></pre>")
				inCodeBlock = false
			} else {
				lang := strings.TrimPrefix(line, "```")
				result.WriteString(fmt.Sprintf("<pre><code class=\"language-%s\">", lang))
				inCodeBlock = true
			}
			continue
		}
		
		if inCodeBlock {
			result.WriteString(escapeHTML(line) + "\n")
			continue
		}
		
		// 标题
		if strings.HasPrefix(line, "# ") {
			result.WriteString("<h1>" + escapeHTML(strings.TrimPrefix(line, "# ")) + "</h1>\n")
			continue
		}
		if strings.HasPrefix(line, "## ") {
			result.WriteString("<h2>" + escapeHTML(strings.TrimPrefix(line, "## ")) + "</h2>\n")
			continue
		}
		if strings.HasPrefix(line, "### ") {
			result.WriteString("<h3>" + escapeHTML(strings.TrimPrefix(line, "### ")) + "</h3>\n")
			continue
		}
		
		// 列表
		if strings.HasPrefix(line, "- ") || strings.HasPrefix(line, "* ") {
			result.WriteString("<li>" + escapeHTML(strings.TrimPrefix(line[2:], " ")) + "</li>\n")
			continue
		}
		
		// 引用
		if strings.HasPrefix(line, "> ") {
			result.WriteString("<blockquote>" + escapeHTML(strings.TrimPrefix(line, "> ")) + "</blockquote>\n")
			continue
		}
		
		// 水平线
		if line == "---" || line == "***" || line == "___" {
			result.WriteString("<hr>\n")
			continue
		}
		
		// 普通段落
		if line != "" {
			result.WriteString("<p>" + processInlineMarkdown(escapeHTML(line)) + "</p>\n")
		}
	}
	
	return result.String()
}

// processInlineMarkdown 处理行内Markdown
func processInlineMarkdown(text string) string {
	// 粗体 **text**
	text = strings.ReplaceAll(text, "**", "")
	// 斜体 *text*
	text = strings.ReplaceAll(text, "*", "")
	// 行内代码 `code`
	// 删除线 ~~text~~
	return text
}

// getLanguageClass 获取语言类名
func getLanguageClass(ext string) string {
	langMap := map[string]string{
		".js":   "javascript",
		".json": "json",
		".xml":  "xml",
		".html": "html",
		".css":  "css",
		".go":   "go",
		".java": "java",
		".py":   "python",
		".c":    "c",
		".cpp":  "cpp",
		".cs":   "csharp",
		".rb":   "ruby",
		".php":  "php",
		".sh":   "bash",
		".yaml": "yaml",
		".yml":  "yaml",
		".sql":  "sql",
	}
	if lang, ok := langMap[ext]; ok {
		return lang
	}
	return "plaintext"
}
