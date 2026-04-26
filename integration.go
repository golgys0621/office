package office

import (
	"fmt"
	"os"
	"path/filepath"
)

// Integration 提供与现有系统集成的便捷方法
type Integration struct {
	handler *DocumentHandler
}

// NewIntegration 创建集成实例
func NewIntegration() *Integration {
	return &Integration{
		handler: NewDocumentHandler(),
	}
}

// SetMammothPath 设置Mammoth路径
func (i *Integration) SetMammothPath(path string) {
	i.handler.SetMammothPath(path)
}

// ConvertFile 转换本地文件
func (i *Integration) ConvertFile(filePath string) (string, error) {
	result, err := i.handler.Convert(filePath)
	if err != nil {
		return "", err
	}
	return result.HTML, nil
}

// ConvertURL 转换远程URL文件
func (i *Integration) ConvertURL(url string) (string, error) {
	result, err := i.handler.ConvertURL(url)
	if err != nil {
		return "", err
	}
	return result.HTML, nil
}

// ConvertAndSave 转换并保存到文件
func (i *Integration) ConvertAndSave(inputPath, outputPath string) error {
	html, err := i.ConvertFile(inputPath)
	if err != nil {
		return err
	}
	return os.WriteFile(outputPath, []byte(html), 0644)
}

// BatchConvert 批量转换
func (i *Integration) BatchConvert(inputs []string, outputDir string) map[string]string {
	results := make(map[string]string)
	
	os.MkdirAll(outputDir, 0755)
	
	for _, input := range inputs {
		ext := filepath.Ext(input)
		output := filepath.Join(outputDir, filepath.Base(input[:len(input)-len(ext)])+".html")
		
		if err := i.ConvertAndSave(input, output); err != nil {
			results[input] = fmt.Sprintf("Error: %v", err)
		} else {
			results[input] = output
		}
	}
	
	return results
}

// GetSupportedFormats 获取支持的格式列表
func (i *Integration) GetSupportedFormats() []string {
	return i.handler.GetSupportedFormats()
}

// IsSupported 检查格式是否支持
func (i *Integration) IsSupported(ext string) bool {
	return i.handler.CanHandle(ext)
}
