package office

import (
	"fmt"
	"sync"
	"time"
)

// Tracer 性能追踪器
type Tracer struct {
	mu      sync.RWMutex
	spans   map[string]*Span
	active  map[string]*Span
	enabled bool
}

// Span 单个追踪跨度
type Span struct {
	Name      string                 `json:"name"`
	StartTime time.Time             `json:"start_time"`
	EndTime   time.Time             `json:"end_time"`
	Duration  time.Duration         `json:"duration"`
	Tags      map[string]interface{} `json:"tags,omitempty"`
	Logs      []SpanLog             `json:"logs,omitempty"`
	ParentID  string               `json:"parent_id,omitempty"`
	ID        string               `json:"id"`
	Status    SpanStatus            `json:"status"`
	Error     string               `json:"error,omitempty"`
}

// SpanLog 跨度日志
type SpanLog struct {
	Time    time.Time              `json:"time"`
	Message string                 `json:"message"`
	Fields  map[string]interface{} `json:"fields,omitempty"`
}

// SpanStatus 跨度状态
type SpanStatus int

const (
	SpanStatusOK       SpanStatus = 0
	SpanStatusError    SpanStatus = 1
	SpanStatusCanceled SpanStatus = 2
)

var globalTracer *Tracer

// 全局追踪器
func init() {
	globalTracer = NewTracer()
}

// NewTracer 创建追踪器
func NewTracer() *Tracer {
	return &Tracer{
		spans:  make(map[string]*Span),
		active: make(map[string]*Span),
	}
}

// Enable 启用追踪
func (t *Tracer) Enable() {
	t.mu.Lock()
	defer t.mu.Unlock()
	t.enabled = true
}

// Disable 禁用追踪
func (t *Tracer) Disable() {
	t.mu.Lock()
	defer t.mu.Unlock()
	t.enabled = false
}

// IsEnabled 检查是否启用
func (t *Tracer) IsEnabled() bool {
	t.mu.RLock()
	defer t.mu.RUnlock()
	return t.enabled
}

// StartSpan 开始一个跨度
func (t *Tracer) StartSpan(name string, opts ...SpanOption) *Span {
	t.mu.Lock()
	defer t.mu.Unlock()

	if !t.enabled {
		return nil
	}

	span := &Span{
		Name:      name,
		StartTime: time.Now(),
		Tags:      make(map[string]interface{}),
		Logs:      make([]SpanLog, 0),
		ID:        generateSpanID(),
		Status:    SpanStatusOK,
	}

	// 应用选项
	for _, opt := range opts {
		opt(span)
	}

	t.active[span.ID] = span
	return span
}

// StopSpan 结束跨度
func (t *Tracer) StopSpan(span *Span, err error) {
	if span == nil {
		return
	}

	t.mu.Lock()
	defer t.mu.Unlock()

	span.EndTime = time.Now()
	span.Duration = span.EndTime.Sub(span.StartTime)

	delete(t.active, span.ID)

	if err != nil {
		span.Status = SpanStatusError
		span.Error = err.Error()
	}

	t.spans[span.ID] = span
}

// Log 记录日志
func (s *Span) Log(message string, fields map[string]interface{}) {
	s.Logs = append(s.Logs, SpanLog{
		Time:    time.Now(),
		Message: message,
		Fields:  fields,
	})
}

// SetTag 设置标签
func (s *Span) SetTag(key string, value interface{}) {
	s.Tags[key] = value
}

// SpanOption 跨度选项
type SpanOption func(*Span)

// WithTag 设置标签
func WithTag(key string, value interface{}) SpanOption {
	return func(s *Span) {
		s.Tags[key] = value
	}
}

// WithTags 设置多个标签
func WithTags(tags map[string]interface{}) SpanOption {
	return func(s *Span) {
		for k, v := range tags {
			s.Tags[k] = v
		}
	}
}

// WithParent 设置父跨度ID
func WithParent(parentID string) SpanOption {
	return func(s *Span) {
		s.ParentID = parentID
	}
}

// GetSpan 获取跨度
func (t *Tracer) GetSpan(id string) *Span {
	t.mu.RLock()
	defer t.mu.RUnlock()
	return t.spans[id]
}

// GetSpans 获取所有跨度
func (t *Tracer) GetSpans() []*Span {
	t.mu.RLock()
	defer t.mu.RUnlock()

	spans := make([]*Span, 0, len(t.spans))
	for _, span := range t.spans {
		spans = append(spans, span)
	}
	return spans
}

// GetActiveSpans 获取活跃跨度
func (t *Tracer) GetActiveSpans() []*Span {
	t.mu.RLock()
	defer t.mu.RUnlock()

	spans := make([]*Span, 0, len(t.active))
	for _, span := range t.active {
		spans = append(spans, span)
	}
	return spans
}

// Clear 清空跨度
func (t *Tracer) Clear() {
	t.mu.Lock()
	defer t.mu.Unlock()
	t.spans = make(map[string]*Span)
	t.active = make(map[string]*Span)
}

// GetStats 获取统计信息
func (t *Tracer) GetStats() TracerStats {
	t.mu.RLock()
	defer t.mu.RUnlock()

	var totalDuration time.Duration
	var errorCount int

	for _, span := range t.spans {
		totalDuration += span.Duration
		if span.Status == SpanStatusError {
			errorCount++
		}
	}

	return TracerStats{
		TotalSpans:    len(t.spans),
		ActiveSpans:   len(t.active),
		TotalDuration: totalDuration,
		ErrorCount:    errorCount,
		SuccessCount:  len(t.spans) - errorCount,
	}
}

// TracerStats 追踪器统计
type TracerStats struct {
	TotalSpans    int           `json:"total_spans"`
	ActiveSpans   int           `json:"active_spans"`
	TotalDuration time.Duration `json:"total_duration"`
	ErrorCount    int           `json:"error_count"`
	SuccessCount  int           `json:"success_count"`
}

// generateSpanID 生成跨度ID
func generateSpanID() string {
	return fmt.Sprintf("%d-%s", time.Now().UnixNano(), randomString(8))
}

// randomString 生成随机字符串
func randomString(n int) string {
	const letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
	b := make([]byte, n)
	for i := range b {
		b[i] = letters[time.Now().UnixNano()%int64(len(letters))]
	}
	return string(b)
}

// 全局追踪函数

// Start 开始追踪
func Start(name string, opts ...SpanOption) *Span {
	return globalTracer.StartSpan(name, opts...)
}

// Stop 结束追踪
func Stop(span *Span, err error) {
	globalTracer.StopSpan(span, err)
}

// EnableTracer 启用追踪
func EnableTracer() {
	globalTracer.Enable()
}

// DisableTracer 禁用追踪
func DisableTracer() {
	globalTracer.Disable()
}

// GetTracerStats 获取追踪统计
func GetTracerStats() TracerStats {
	return globalTracer.GetStats()
}

// TraceFunc 追踪函数执行
func TraceFunc(name string, fn func() error) error {
	span := Start(name)
	err := fn()
	Stop(span, err)
	return err
}

// TraceFuncResult 追踪函数执行并返回结果
func TraceFuncResult(name string, fn func() (interface{}, error)) (interface{}, error) {
	span := Start(name)
	result, err := fn()
	Stop(span, err)
	return result, err
}

// TraceOperation 追踪操作
type TraceOperation struct {
	tracer *Tracer
	span   *Span
	name   string
	start  time.Time
}

// StartTrace 开始追踪
func StartTrace(name string) *TraceOperation {
	span := globalTracer.StartSpan(name)
	return &TraceOperation{
		tracer: globalTracer,
		span:   span,
		name:   name,
		start:  time.Now(),
	}
}

// Log 记录日志
func (t *TraceOperation) Log(message string, fields map[string]interface{}) {
	if t.span != nil {
		t.span.Log(message, fields)
	}
}

// SetTag 设置标签
func (t *TraceOperation) SetTag(key string, value interface{}) {
	if t.span != nil {
		t.span.SetTag(key, value)
	}
}

// End 结束追踪
func (t *TraceOperation) End(err error) {
	if t.span != nil {
		t.span.EndTime = time.Now()
		t.span.Duration = t.span.EndTime.Sub(t.start)
		if err != nil {
			t.span.Status = SpanStatusError
			t.span.Error = err.Error()
		}
		globalTracer.StopSpan(t.span, err)
	}
}

// Finish 结束追踪（无错误）
func (t *TraceOperation) Finish() {
	t.End(nil)
}

// Panic 追踪panic并恢复
func (t *TraceOperation) Panic() {
	if r := recover(); r != nil {
		if t.span != nil {
			t.span.Status = SpanStatusError
			t.span.Error = fmt.Sprintf("panic: %v", r)
		}
		t.End(fmt.Errorf("%v", r))
		panic(r)
	}
}
