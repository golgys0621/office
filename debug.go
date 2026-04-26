package office

import (
	"encoding/json"
	"fmt"
	"os"
	"runtime"
	"strings"
	"sync"
	"time"
)

// DebugLevel 调试级别
type DebugLevel int

const (
	DebugLevelOff     DebugLevel = 0
	DebugLevelError   DebugLevel = 1
	DebugLevelWarn    DebugLevel = 2
	DebugLevelInfo    DebugLevel = 3
	DebugLevelDebug   DebugLevel = 4
	DebugLevelTrace   DebugLevel = 5
)

// String 转换为字符串
func (l DebugLevel) String() string {
	switch l {
	case DebugLevelOff:
		return "OFF"
	case DebugLevelError:
		return "ERROR"
	case DebugLevelWarn:
		return "WARN"
	case DebugLevelInfo:
		return "INFO"
	case DebugLevelDebug:
		return "DEBUG"
	case DebugLevelTrace:
		return "TRACE"
	default:
		return "UNKNOWN"
	}
}

// ParseDebugLevel 解析调试级别字符串
func ParseDebugLevel(s string) DebugLevel {
	switch strings.ToUpper(s) {
	case "OFF", "0":
		return DebugLevelOff
	case "ERROR", "1":
		return DebugLevelError
	case "WARN", "WARNING", "2":
		return DebugLevelWarn
	case "INFO", "3":
		return DebugLevelInfo
	case "DEBUG", "4":
		return DebugLevelDebug
	case "TRACE", "5":
		return DebugLevelTrace
	default:
		return DebugLevelInfo
	}
}

// LogEntry 日志条目
type LogEntry struct {
	Time      time.Time              `json:"time"`
	Level     DebugLevel             `json:"level"`
	Category  string                 `json:"category"`
	Message   string                 `json:"message"`
	Details   map[string]interface{} `json:"details,omitempty"`
	Stack     string                 `json:"stack,omitempty"`
	Duration  time.Duration          `json:"duration,omitempty"`
	RequestID string                 `json:"request_id,omitempty"`
	Error     string                 `json:"error,omitempty"`
}

// DebugConfig 调试配置
type DebugConfig struct {
	Level          DebugLevel   // 日志级别
	EnableFile     bool         // 启用文件日志
	EnableConsole  bool         // 启用控制台输出
	LogFilePath    string       // 日志文件路径
	MaxFileSize    int64        // 单个日志文件最大大小 (MB)
	MaxBackupFiles int          // 保留的日志文件数量
	EnableColors   bool         // 启用彩色输出
	EnableJSON     bool         // 输出JSON格式
	EnableStack    bool         // 记录错误堆栈
	EnablePerf     bool         // 启用性能监控
	BufferSize     int          // 日志缓冲区大小
}

// DefaultDebugConfig 默认调试配置
var DefaultDebugConfig = &DebugConfig{
	Level:          DebugLevelInfo,
	EnableConsole:  true,
	EnableFile:     false,
	LogFilePath:    "logs/office-debug.log",
	MaxFileSize:    10,
	MaxBackupFiles: 5,
	EnableColors:   true,
	EnableJSON:     false,
	EnableStack:    true,
	EnablePerf:     true,
	BufferSize:     100,
}

// Logger 调试日志器
type Logger struct {
	mu       sync.RWMutex
	config   *DebugConfig
	entries  []LogEntry // 内存缓冲区
	writer   *os.File
	perfData map[string]*perfRecord
}

// perfRecord 性能记录
type perfRecord struct {
	Count    int
	Total    time.Duration
	Min      time.Duration
	Max      time.Duration
	Avg      time.Duration
	LastTime time.Time
}

var (
	defaultLogger *Logger
	once          sync.Once
)

// GetLogger 获取默认日志器
func GetLogger() *Logger {
	once.Do(func() {
		defaultLogger = NewLogger(DefaultDebugConfig)
	})
	return defaultLogger
}

// NewLogger 创建日志器
func NewLogger(config *DebugConfig) *Logger {
	l := &Logger{
		config:   config,
		entries:  make([]LogEntry, 0, config.BufferSize),
		perfData: make(map[string]*perfRecord),
	}

	if config.EnableFile {
		if err := l.initFileWriter(); err != nil {
			fmt.Printf("Failed to init file writer: %v\n", err)
		}
	}

	return l
}

// initFileWriter 初始化文件写入器
func (l *Logger) initFileWriter() error {
	// 确保目录存在
	dir := l.config.LogFilePath[:strings.LastIndex(l.config.LogFilePath, "/")]
	if dir != "" {
		os.MkdirAll(dir, 0755)
	}

	file, err := os.OpenFile(l.config.LogFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if err != nil {
		return err
	}
	l.writer = file
	return nil
}

// Close 关闭日志器
func (l *Logger) Close() error {
	if l.writer != nil {
		return l.writer.Close()
	}
	return nil
}

// SetLevel 设置日志级别
func (l *Logger) SetLevel(level DebugLevel) {
	l.mu.Lock()
	defer l.mu.Unlock()
	l.config.Level = level
}

// SetConfig 设置配置
func (l *Logger) SetConfig(config *DebugConfig) {
	l.mu.Lock()
	defer l.mu.Unlock()
	l.config = config
}

// Log 记录日志
func (l *Logger) Log(level DebugLevel, category, message string, details map[string]interface{}) {
	if level > l.config.Level {
		return
	}

	entry := LogEntry{
		Time:     time.Now(),
		Level:    level,
		Category: category,
		Message:  message,
		Details:  details,
	}

	// 记录堆栈信息
	if level <= DebugLevelError && l.config.EnableStack {
		entry.Stack = l.getStack()
	}

	l.mu.Lock()
	l.entries = append(l.entries, entry)
	if len(l.entries) > l.config.BufferSize {
		l.entries = l.entries[1:]
	}
	l.mu.Unlock()

	// 输出到控制台
	if l.config.EnableConsole {
		l.writeToConsole(entry)
	}

	// 输出到文件
	if l.config.EnableFile && l.writer != nil {
		l.writeToFile(entry)
	}
}

// Error 记录错误
func (l *Logger) Error(category, message string, details map[string]interface{}) {
	l.Log(DebugLevelError, category, message, details)
}

// Warn 记录警告
func (l *Logger) Warn(category, message string, details map[string]interface{}) {
	l.Log(DebugLevelWarn, category, message, details)
}

// Info 记录信息
func (l *Logger) Info(category, message string, details map[string]interface{}) {
	l.Log(DebugLevelInfo, category, message, details)
}

// Debug 记录调试信息
func (l *Logger) Debug(category, message string, details map[string]interface{}) {
	l.Log(DebugLevelDebug, category, message, details)
}

// Trace 记录追踪信息
func (l *Logger) Trace(category, message string, details map[string]interface{}) {
	l.Log(DebugLevelTrace, category, message, details)
}

// getStack 获取堆栈信息
func (l *Logger) getStack() string {
	buf := make([]byte, 4096)
	n := runtime.Stack(buf, false)
	return string(buf[:n])
}

// writeToConsole 输出到控制台
func (l *Logger) writeToConsole(entry LogEntry) {
	var color string
	switch entry.Level {
	case DebugLevelError:
		color = "\033[31m" // 红色
	case DebugLevelWarn:
		color = "\033[33m" // 黄色
	case DebugLevelInfo:
		color = "\033[32m" // 绿色
	case DebugLevelDebug:
		color = "\033[36m" // 青色
	case DebugLevelTrace:
		color = "\033[90m" // 灰色
	default:
		color = "\033[0m"
	}
	reset := "\033[0m"

	if l.config.EnableJSON {
		data, _ := json.Marshal(entry)
		fmt.Println(string(data))
	} else {
		levelStr := fmt.Sprintf("[%s]", entry.Level.String())
		timeStr := entry.Time.Format("2006-01-02 15:04:05.000")
		catStr := fmt.Sprintf("[%s]", entry.Category)

		if l.config.EnableColors {
			fmt.Printf("%s%s %s%s %s %s%s\n%s",
				color, timeStr, levelStr, catStr, entry.Message, reset, getDetailsString(entry.Details), entry.Stack)
		} else {
			fmt.Printf("%s %s %s %s\n%s%s\n",
				timeStr, levelStr, catStr, entry.Message, getDetailsString(entry.Details), entry.Stack)
		}
	}
}

// writeToFile 输出到文件
func (l *Logger) writeToFile(entry LogEntry) {
	l.mu.Lock()
	defer l.mu.Unlock()

	data, err := json.Marshal(entry)
	if err != nil {
		return
	}

	l.writer.Write(append(data, '\n'))

	// 检查文件大小
	if l.config.MaxFileSize > 0 {
		if info, err := l.writer.Stat(); err == nil && info.Size() > l.config.MaxFileSize*1024*1024 {
			l.rotateFile()
		}
	}
}

// rotateFile 轮转日志文件
func (l *Logger) rotateFile() {
	l.writer.Close()

	// 移动旧文件
	oldPath := l.config.LogFilePath
	newPath := fmt.Sprintf("%s.%s", oldPath, time.Now().Format("20060102150405"))
	os.Rename(oldPath, newPath)

	// 重新打开文件
	file, _ := os.OpenFile(l.config.LogFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	l.writer = file
}

// GetEntries 获取日志条目
func (l *Logger) GetEntries(level DebugLevel, category string, limit int) []LogEntry {
	l.mu.RLock()
	defer l.mu.RUnlock()

	var result []LogEntry
	for i := len(l.entries) - 1; i >= 0 && len(result) < limit; i-- {
		entry := l.entries[i]
		if level > 0 && entry.Level < level {
			continue
		}
		if category != "" && entry.Category != category {
			continue
		}
		result = append(result, entry)
	}
	return result
}

// ClearEntries 清除日志条目
func (l *Logger) ClearEntries() {
	l.mu.Lock()
	defer l.mu.Unlock()
	l.entries = make([]LogEntry, 0, l.config.BufferSize)
}

// RecordPerf 记录性能数据
func (l *Logger) RecordPerf(name string, duration time.Duration) {
	if !l.config.EnablePerf {
		return
	}

	l.mu.Lock()
	defer l.mu.Unlock()

	record, exists := l.perfData[name]
	if !exists {
		record = &perfRecord{
			Min: duration,
			Max: duration,
		}
		l.perfData[name] = record
	}

	record.Count++
	record.Total += duration
	record.Avg = record.Total / time.Duration(record.Count)
	record.LastTime = time.Now()

	if duration < record.Min {
		record.Min = duration
	}
	if duration > record.Max {
		record.Max = duration
	}
}

// GetPerfStats 获取性能统计
func (l *Logger) GetPerfStats(name string) *PerfStats {
	l.mu.RLock()
	defer l.mu.RUnlock()

	if record, exists := l.perfData[name]; exists {
		return &PerfStats{
			Name:  name,
			Count: record.Count,
			Total: record.Total,
			Min:   record.Min,
			Max:   record.Max,
			Avg:   record.Avg,
		}
	}
	return nil
}

// GetAllPerfStats 获取所有性能统计
func (l *Logger) GetAllPerfStats() []*PerfStats {
	l.mu.RLock()
	defer l.mu.RUnlock()

	var stats []*PerfStats
	for name, record := range l.perfData {
		stats = append(stats, &PerfStats{
			Name:  name,
			Count: record.Count,
			Total: record.Total,
			Min:   record.Min,
			Max:   record.Max,
			Avg:   record.Avg,
		})
	}
	return stats
}

// PerfStats 性能统计
type PerfStats struct {
	Name  string        `json:"name"`
	Count int           `json:"count"`
	Total time.Duration `json:"total"`
	Min   time.Duration `json:"min"`
	Max   time.Duration `json:"max"`
	Avg   time.Duration `json:"avg"`
}

// getDetailsString 格式化详情
func getDetailsString(details map[string]interface{}) string {
	if details == nil {
		return ""
	}
	data, err := json.Marshal(details)
	if err != nil {
		return ""
	}
	return " " + string(data)
}

// 便捷函数

// Error 记录错误
func Error(category, message string, details map[string]interface{}) {
	GetLogger().Error(category, message, details)
}

// Warn 记录警告
func Warn(category, message string, details map[string]interface{}) {
	GetLogger().Warn(category, message, details)
}

// Info 记录信息
func Info(category, message string, details map[string]interface{}) {
	GetLogger().Info(category, message, details)
}

// Debug 记录调试信息
func Debug(category, message string, details map[string]interface{}) {
	GetLogger().Debug(category, message, details)
}

// Trace 记录追踪信息
func Trace(category, message string, details map[string]interface{}) {
	GetLogger().Trace(category, message, details)
}

// SetLogLevel 设置日志级别
func SetLogLevel(level DebugLevel) {
	GetLogger().SetLevel(level)
}

// SetLogLevelFromEnv 从环境变量设置日志级别
func SetLogLevelFromEnv() {
	if level := os.Getenv("OFFICE_DEBUG"); level != "" {
		GetLogger().SetLevel(ParseDebugLevel(level))
	}
}
