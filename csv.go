package office

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// csvHandler CSV处理器
type csvHandler struct{}

// newCsvHandler 创建CSV处理器
func newCsvHandler() *csvHandler {
	return &csvHandler{}
}

// Name 返回处理器名称
func (h *csvHandler) Name() string {
	return "csv"
}

// Extensions 返回支持的扩展名
func (h *csvHandler) Extensions() []string {
	return []string{"csv", "tsv"}
}

// CanHandle 检查是否支持该格式
func (h *csvHandler) CanHandle(ext string) bool {
	ext = strings.ToLower(ext)
	for _, e := range h.Extensions() {
		if ext == e {
			return true
		}
	}
	return false
}

// ToHTML 转换为HTML
func (h *csvHandler) ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, NewError(ErrFileNotFound, "文件不存在: "+filePath, err)
	}

	file, err := os.Open(filePath)
	if err != nil {
		return nil, NewError(ErrConversionFailed, "打开文件失败: "+err.Error(), err)
	}
	defer file.Close()

	ext := strings.ToLower(filepath.Ext(filePath))
	delimiter := ','
	if ext == ".tsv" {
		delimiter = '\t'
	}

	reader := csv.NewReader(file)
	reader.Comma = delimiter
	reader.LazyQuotes = true

	records, err := reader.ReadAll()
	if err != nil {
		return nil, NewError(ErrConversionFailed, "解析CSV失败: "+err.Error(), err)
	}

	html := h.generateTable(records, filepath.Base(filePath), opts)

	return &HTMLResult{
		HTML: html,
		Metadata: &Metadata{
			Title:       filepath.Base(filePath),
			FileSize:    getFileSize(filePath),
			Modified:    getFileModTime(filePath),
		},
		Format: TypeCSV,
	}, nil
}

// generateTable 生成HTML表格
func (h *csvHandler) generateTable(records [][]string, fileName string, opts *ConversionOptions) string {
	if len(records) == 0 {
		return `<div class="csv-empty">CSV文件为空</div>`
	}

	var html strings.Builder

	html.WriteString(`<div class="csv-container">
  <div class="csv-header">
    <span class="csv-title">` + escapeHTML(fileName) + `</span>
    <span class="csv-info">` + fmt.Sprintf("%d行 x %d列", len(records), len(records[0])) + `</span>
  </div>
  <div class="table-wrapper">
    <table>
`)

	for rowIdx, row := range records {
		tag := "td"
		class := ""
		if rowIdx == 0 {
			tag = "th"
			class = " class=\"header-row\""
		}
		
		html.WriteString("      <tr>\n")
		for _, cell := range row {
			html.WriteString(fmt.Sprintf("        <%s%v>%s</%s>\n", tag, class, escapeHTML(cell), tag))
		}
		html.WriteString("      </tr>\n")
	}

	html.WriteString(`    </table>
  </div>
</div>`)

	css := h.getCSS()

	return fmt.Sprintf(`<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s</title>
<style>%s</style>
</head>
<body>
%s
</body>
</html>`, fileName, css, html.String())
}

// getCSS 获取CSS样式
func (h *csvHandler) getCSS() string {
	return `
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Microsoft YaHei', Arial, sans-serif; background: #f3f3f3; }
.csv-container { max-width: 1200px; margin: 0 auto; }
.csv-header { 
    background: #217346; 
    color: white; 
    padding: 12px 20px; 
    display: flex; 
    justify-content: space-between;
}
.csv-title { font-size: 14px; }
.csv-info { font-size: 12px; color: rgba(255,255,255,0.8); }
.table-wrapper { overflow-x: auto; padding: 16px; background: white; }
table { 
    border-collapse: collapse; 
    width: 100%%; 
    font-size: 13px;
}
th, td { 
    border: 1px solid #d4d4d4; 
    padding: 8px 12px; 
    text-align: left;
    max-width: 300px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
th { background: #e8e8e8; font-weight: bold; position: sticky; top: 0; }
tr:hover { background: #f5f5f5; }
tr:nth-child(even) { background: #fafafa; }
.header-row { background: #e8e8e8; font-weight: bold; }
.csv-empty { padding: 40px; text-align: center; color: #888; }
`
}

// GetMetadata 获取文档元信息
func (h *csvHandler) GetMetadata(filePath string) (*Metadata, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, NewError(ErrConversionFailed, "打开文件失败: "+err.Error(), err)
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		return nil, NewError(ErrConversionFailed, "解析CSV失败: "+err.Error(), err)
	}

	rows := len(records)
	_ = rows // 保留以备后用
	// cols := 0
	// if rows > 0 {
	//     cols = len(records[0])
	// }

	return &Metadata{
		Title:    filepath.Base(filePath),
		FileSize: getFileSize(filePath),
		Modified: getFileModTime(filePath),
	}, nil
}

// ReadCSV 读取CSV文件
func ReadCSV(filePath string) ([][]string, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	reader := csv.NewReader(bufio.NewReader(file))
	return reader.ReadAll()
}

// WriteCSV 写入CSV文件
func WriteCSV(filePath string, records [][]string) error {
	file, err := os.Create(filePath)
	if err != nil {
		return err
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	return writer.WriteAll(records)
}
