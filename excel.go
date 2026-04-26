package office

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// excelHandler Excel文档处理器
type excelHandler struct{}

// newExcelHandler 创建Excel处理器
func newExcelHandler() *excelHandler {
	return &excelHandler{}
}

// Name 返回处理器名称
func (h *excelHandler) Name() string {
	return "excel"
}

// Extensions 返回支持的扩展名
func (h *excelHandler) Extensions() []string {
	return []string{"xlsx", "xls", "xlsm", "et", "ett", "csv"}
}

// CanHandle 检查是否支持该格式
func (h *excelHandler) CanHandle(ext string) bool {
	ext = strings.ToLower(ext)
	for _, e := range h.Extensions() {
		if ext == e {
			return true
		}
	}
	return false
}

// ToHTML 转换为HTML
func (h *excelHandler) ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, NewError(ErrFileNotFound, "文件不存在: "+filePath, err)
	}

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, NewError(ErrConversionFailed, "打开Excel文件失败: "+err.Error(), err)
	}
	defer f.Close()

	// 生成HTML
	html := h.generateHTML(f, filepath.Base(filePath), opts)

	// 获取元信息
	sheets := f.GetSheetList()
	metadata := &Metadata{
		Title:       filepath.Base(filePath),
		SheetCount:  len(sheets),
		FileSize:    getFileSize(filePath),
		Modified:    getFileModTime(filePath),
	}

	return &HTMLResult{
		HTML:     html,
		Metadata: metadata,
		Format:   TypeXLSX,
	}, nil
}

// generateHTML 生成HTML
func (h *excelHandler) generateHTML(f *excelize.File, fileName string, opts *ConversionOptions) string {
	sheets := f.GetSheetList()
	
	var tabsHTML, contentHTML strings.Builder

	// 生成工作表标签
	tabsHTML.WriteString(`<div class="excel-tabs">`)
	for i, sheet := range sheets {
		active := ""
		if i == 0 {
			active = " active"
		}
		tabsHTML.WriteString(fmt.Sprintf(
			`<button class="tab-btn%s" data-sheet="%d" onclick="showSheet(%d)">%s</button>`,
			active, i, i, escapeHTML(sheet)))
	}
	tabsHTML.WriteString(`</div>`)

	// 生成工作表内容
	contentHTML.WriteString(`<div class="excel-sheets">`)
	for i, sheet := range sheets {
		display := "none"
		if i == 0 {
			display = "block"
		}
		contentHTML.WriteString(fmt.Sprintf(`<div class="sheet-content" id="sheet-%d" style="display:%s">`, i, display))
		contentHTML.WriteString(h.sheetToHTML(f, sheet))
		contentHTML.WriteString(`</div>`)
	}
	contentHTML.WriteString(`</div>`)

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
<div class="excel-container">
  <div class="excel-header">%s</div>
  %s
</div>
<script>%s</script>
</body>
</html>`, fileName, css, tabsHTML.String(), contentHTML.String(), js)
}

// sheetToHTML 将工作表转换为HTML表格
func (h *excelHandler) sheetToHTML(f *excelize.File, sheetName string) string {
	var html strings.Builder

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return `<p class="error">读取工作表失败</p>`
	}

	if len(rows) == 0 {
		return `<p class="empty">工作表为空</p>`
	}

	html.WriteString(`<div class="table-wrapper"><table>`)

	for rowIdx, row := range rows {
		html.WriteString("<tr>")
		for colIdx, cell := range row {
			cellRef := toCellRef(rowIdx, colIdx)
			
			// 获取单元格样式
			style, _ := f.GetCellStyle(sheetName, cellRef)
			styleClass := getExcelStyleClass(style)

			// 处理单元格内容
			content := escapeHTML(cell)
			if content == "" {
				content = "&nbsp;"
			}

			tag := "td"
			if rowIdx == 0 {
				tag = "th"
			}

			attrs := styleClass

			html.WriteString(fmt.Sprintf("<%s class=\"%s\">%s</%s>", tag, attrs, content, tag))
		}
		html.WriteString("</tr>")
	}

	html.WriteString("</table></div>")
	return html.String()
}

// getCSS 获取Excel CSS样式
func (h *excelHandler) getCSS(opts *ConversionOptions) string {
	maxWidth := 1200
	if opts != nil && opts.MaxWidth > 0 {
		maxWidth = opts.MaxWidth
	}

	return fmt.Sprintf(`
* { box-sizing: border-box; margin: 0; padding: 0; }
body { 
    font-family: 'Microsoft YaHei', Arial, sans-serif; 
    background: #f3f3f3; 
    padding: 20px;
}
.excel-container { 
    max-width: %dpx; 
    margin: 0 auto; 
    background: white; 
    border-radius: 8px; 
    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    overflow: hidden;
}
.excel-header { 
    background: #217346; 
    color: white; 
    padding: 12px 20px; 
    font-size: 16px;
    font-weight: bold;
}
.excel-tabs { 
    display: flex; 
    background: #e8e8e8; 
    border-bottom: 1px solid #d4d4d4;
    overflow-x: auto;
}
.tab-btn { 
    padding: 10px 20px; 
    border: none; 
    background: transparent; 
    cursor: pointer; 
    border-bottom: 2px solid transparent;
    white-space: nowrap;
}
.tab-btn:hover { background: #d4d4d4; }
.tab-btn.active { background: white; border-bottom-color: #217346; }
.sheet-content { overflow-x: auto; padding: 16px; }
.table-wrapper { overflow-x: auto; }
table { 
    border-collapse: collapse; 
    width: 100%%; 
    min-width: 600px;
}
th, td { 
    border: 1px solid #d4d4d4; 
    padding: 8px 12px; 
    text-align: left;
    white-space: nowrap;
    max-width: 300px;
    overflow: hidden;
    text-overflow: ellipsis;
}
th { background: #f5f5f5; font-weight: bold; }
tr:nth-child(even) { background: #fafafa; }
tr:hover { background: #f0f7f0; }
.cell-bold { font-weight: bold; }
.cell-italic { font-style: italic; }
.cell-center { text-align: center; }
.cell-right { text-align: right; }
.error { color: #dc3545; padding: 20px; text-align: center; }
.empty { color: #888; padding: 40px; text-align: center; }
`, maxWidth)
}

// getJS 获取JavaScript代码
func (h *excelHandler) getJS() string {
	return `
function showSheet(index) {
    document.querySelectorAll('.sheet-content').forEach((el, i) => {
        el.style.display = i === index ? 'block' : 'none';
    });
    document.querySelectorAll('.tab-btn').forEach((el, i) => {
        el.classList.toggle('active', i === index);
    });
}
`
}

// GetMetadata 获取文档元信息
func (h *excelHandler) GetMetadata(filePath string) (*Metadata, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, NewError(ErrConversionFailed, "打开Excel文件失败: "+err.Error(), err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	
	return &Metadata{
		Title:       filepath.Base(filePath),
		SheetCount:  len(sheets),
		FileSize:    getFileSize(filePath),
		Modified:    getFileModTime(filePath),
	}, nil
}

// ============ 辅助函数 ============

// toCellRef 转换为Excel单元格引用
func toCellRef(row, col int) string {
	colStr := ""
	for col >= 0 {
		colStr = string(rune('A'+col%26)) + colStr
		col = col/26 - 1
	}
	return fmt.Sprintf("%s%d", colStr, row+1)
}

// getMergedSpan 获取合并单元格跨度
func getMergedSpan(cellRef string, mergedCells []string) (rowspan, colspan int) {
	// 简化实现
	return 1, 1
}

// getExcelStyleClass 获取Excel样式类名
// styleIdx 是从GetCellStyle返回的样式索引
func getExcelStyleClass(styleIdx int) string {
	if styleIdx < 0 {
		return ""
	}
	
	// 简化实现 - 实际应该使用 excelize 的样式相关API
	// 这里返回空字符串，由后续版本完善
	return ""
}

// getFileSize 获取文件大小
func getFileSize(path string) int64 {
	info, _ := os.Stat(path)
	if info != nil {
		return info.Size()
	}
	return 0
}

// getFileModTime 获取文件修改时间
func getFileModTime(path string) (t time.Time) {
	info, _ := os.Stat(path)
	if info != nil {
		return info.ModTime()
	}
	return
}

// escapeHTML HTML转义
func escapeHTML(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	return s
}
