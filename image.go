package office

import (
	"encoding/base64"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// imageHandler 图片处理器
type imageHandler struct{}

// newImageHandler 创建图片处理器
func newImageHandler() *imageHandler {
	return &imageHandler{}
}

// Name 返回处理器名称
func (h *imageHandler) Name() string {
	return "image"
}

// Extensions 返回支持的扩展名
func (h *imageHandler) Extensions() []string {
	return []string{"jpg", "jpeg", "png", "gif", "bmp", "webp", "svg", "ico"}
}

// CanHandle 检查是否支持该格式
func (h *imageHandler) CanHandle(ext string) bool {
	ext = strings.ToLower(ext)
	for _, e := range h.Extensions() {
		if ext == e {
			return true
		}
	}
	return false
}

// ToHTML 转换为HTML
func (h *imageHandler) ToHTML(filePath string, opts *ConversionOptions) (*HTMLResult, error) {
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return nil, NewError(ErrFileNotFound, "文件不存在: "+filePath, err)
	}

	// 读取图片文件
	imageData, err := os.ReadFile(filePath)
	if err != nil {
		return nil, NewError(ErrConversionFailed, "读取图片失败: "+err.Error(), err)
	}

	// 获取MIME类型
	mimeType := getImageMimeType(filePath)
	
	// 获取图片尺寸
	width, height := getImageDimensions(filePath)

	// 包装HTML
	html := h.wrapHTML(filePath, filepath.Base(filePath), opts, mimeType, width, height)

	return &HTMLResult{
		HTML: html,
		Metadata: &Metadata{
			Title:    filepath.Base(filePath),
			FileSize: int64(len(imageData)),
			Modified: getFileModTime(filePath),
		},
		Images: map[string][]byte{
			filepath.Base(filePath): imageData,
		},
		Format: TypeImage,
	}, nil
}

// wrapHTML 包装HTML
func (h *imageHandler) wrapHTML(filePath, fileName string, opts *ConversionOptions, mimeType string, width, height int) string {
	// 读取图片数据
	imageData, _ := os.ReadFile(filePath)
	dataURL := fmt.Sprintf("data:%s;base64,%s", mimeType, base64.StdEncoding.EncodeToString(imageData))

	css := h.getCSS(opts)

	return fmt.Sprintf(`<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s</title>
<style>%s</style>
</head>
<body>
<div class="image-container">
  <div class="image-header">
    <span class="image-title">%s</span>
    <span class="image-info">%dx%d</span>
  </div>
  <div class="image-viewer">
    <img id="image" src="%s" alt="%s">
  </div>
  <div class="image-controls">
    <button onclick="zoomIn()">放大</button>
    <button onclick="zoomOut()">缩小</button>
    <button onclick="resetZoom()">重置</button>
    <span id="zoom-level">100%%</span>
  </div>
</div>
<script>
let scale = 1;
const img = document.getElementById('image');
function updateZoom() {
    img.style.transform = 'scale(' + scale + ')';
    document.getElementById('zoom-level').textContent = Math.round(scale * 100) + '%%';
}
function zoomIn() { scale = Math.min(scale + 0.25, 5); updateZoom(); }
function zoomOut() { scale = Math.max(scale - 0.25, 0.25); updateZoom(); }
function resetZoom() { scale = 1; updateZoom(); }
document.addEventListener('wheel', (e) => {
    if (e.ctrlKey) {
        e.preventDefault();
        if (e.deltaY < 0) zoomIn(); else zoomOut();
    }
});
</script>
</body>
</html>`, fileName, css, fileName, width, height, dataURL, fileName)
}

// getCSS 获取CSS样式
func (h *imageHandler) getCSS(opts *ConversionOptions) string {
	maxWidth := 1200
	if opts != nil && opts.MaxWidth > 0 {
		maxWidth = opts.MaxWidth
	}

	return fmt.Sprintf(`
* { box-sizing: border-box; margin: 0; padding: 0; }
body { background: #1e1e1e; font-family: Arial, sans-serif; }
.image-container { max-width: %dpx; margin: 0 auto; }
.image-header { background: #333; color: white; padding: 12px 20px; display: flex; justify-content: space-between; }
.image-title { font-size: 14px; }
.image-info { font-size: 12px; color: #aaa; }
.image-viewer { 
    background: #2d2d2d; 
    min-height: 400px; 
    display: flex; 
    align-items: center; 
    justify-content: center;
    overflow: auto;
    padding: 20px;
}
#image { 
    max-width: 100%%; 
    transition: transform 0.2s; 
    cursor: zoom-in;
}
.image-controls {
    background: #333;
    padding: 12px 20px;
    display: flex;
    gap: 10px;
    align-items: center;
}
button {
    padding: 8px 16px;
    background: #4a4a4a;
    color: white;
    border: none;
    border-radius: 4px;
    cursor: pointer;
}
button:hover { background: #5a5a5a; }
#zoom-level { color: #aaa; margin-left: 20px; }
`, maxWidth)
}

// GetMetadata 获取文档元信息
func (h *imageHandler) GetMetadata(filePath string) (*Metadata, error) {
	imageData, _ := os.ReadFile(filePath)
	return &Metadata{
		Title:    filepath.Base(filePath),
		FileSize: int64(len(imageData)),
		Modified: getFileModTime(filePath),
	}, nil
}

// getImageMimeType 获取图片MIME类型
func getImageMimeType(filePath string) string {
	ext := strings.ToLower(filepath.Ext(filePath))
	mimeTypes := map[string]string{
		".jpg":  "image/jpeg",
		".jpeg": "image/jpeg",
		".png":  "image/png",
		".gif":  "image/gif",
		".bmp":  "image/bmp",
		".webp": "image/webp",
		".svg":  "image/svg+xml",
		".ico":  "image/x-icon",
	}
	if mime, ok := mimeTypes[ext]; ok {
		return mime
	}
	return "application/octet-stream"
}

// getImageDimensions 获取图片尺寸（简化实现）
func getImageDimensions(filePath string) (int, int) {
	// 简化实现，实际应该解析图片头
	data, _ := os.ReadFile(filePath)
	
	// PNG
	if len(data) > 24 && data[0] == 0x89 && data[1] == 'P' && data[2] == 'N' && data[3] == 'G' {
		width := int(data[16])<<24 | int(data[17])<<16 | int(data[18])<<8 | int(data[19])
		height := int(data[20])<<24 | int(data[21])<<16 | int(data[22])<<8 | int(data[23])
		return width, height
	}
	
	// JPEG
	if len(data) > 2 && data[0] == 0xFF && data[1] == 0xD8 {
		i := 2
		for i < len(data)-1 {
			if data[i] != 0xFF {
				i++
				continue
			}
			marker := data[i+1]
			if marker == 0xC0 || marker == 0xC2 {
				height := int(data[i+5])<<8 | int(data[i+6])
				width := int(data[i+7])<<8 | int(data[i+8])
				return width, height
			}
			length := int(data[i+2])<<8 | int(data[i+3])
			i += 2 + length
		}
	}
	
	// GIF
	if len(data) > 10 && data[0] == 'G' && data[1] == 'I' && data[2] == 'F' {
		width := int(data[6]) | int(data[7])<<8
		height := int(data[8]) | int(data[9])<<8
		return width, height
	}
	
	// 默认尺寸
	return 800, 600
}
