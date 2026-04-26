package office

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"runtime"
	"sort"
	"strings"
)

// DiagnosticInfo 诊断信息
type DiagnosticInfo struct {
	SystemInfo    SystemInfo           `json:"system_info"`
	GoInfo        GoInfo              `json:"go_info"`
	NodeInfo      NodeInfo            `json:"node_info"`
	Environment   map[string]string   `json:"environment"`
	Dependencies  []DependencyInfo    `json:"dependencies"`
	RecentErrors  []LogEntry          `json:"recent_errors,omitempty"`
	PerfStats     []*PerfStats        `json:"perf_stats,omitempty"`
	TracerStats   TracerStats         `json:"tracer_stats,omitempty"`
}

// SystemInfo 系统信息
type SystemInfo struct {
	OS           string `json:"os"`
	Arch         string `json:"arch"`
	NumCPU       int    `json:"num_cpu"`
	GoMaxProcs   int    `json:"go_max_procs"`
	TotalMemory  uint64 `json:"total_memory"`
	FreeMemory   uint64 `json:"free_memory"`
	Hostname     string `json:"hostname"`
	CurrentDir   string `json:"current_dir"`
	TempDir      string `json:"temp_dir"`
}

// GoInfo Go信息
type GoInfo struct {
	Version      string `json:"version"`
	GOOS         string `json:"goos"`
	GOARCH       string `json:"goarch"`
	Compiler     string `json:"compiler"`
	BuildSettings map[string]string `json:"build_settings,omitempty"`
}

// NodeInfo Node.js信息
type NodeInfo struct {
	Installed    bool   `json:"installed"`
	Version      string `json:"version"`
	Path         string `json:"path"`
	NpmVersion   string `json:"npm_version,omitempty"`
}

// DependencyInfo 依赖信息
type DependencyInfo struct {
	Name     string `json:"name"`
	Version  string `json:"version"`
	Status   string `json:"status"`
	Location string `json:"location,omitempty"`
}

// CollectDiagnostics 收集诊断信息
func CollectDiagnostics() *DiagnosticInfo {
	info := &DiagnosticInfo{
		SystemInfo:   collectSystemInfo(),
		GoInfo:       collectGoInfo(),
		NodeInfo:     collectNodeInfo(),
		Environment:  collectEnvironment(),
		Dependencies: collectDependencies(),
	}

	// 获取最近错误
	logger := GetLogger()
	info.RecentErrors = logger.GetEntries(DebugLevelError, "", 20)
	info.PerfStats = logger.GetAllPerfStats()
	info.TracerStats = GetTracerStats()

	return info
}

// collectSystemInfo 收集系统信息
func collectSystemInfo() SystemInfo {
	info := SystemInfo{
		OS:         runtime.GOOS,
		Arch:       runtime.GOARCH,
		NumCPU:     runtime.NumCPU(),
		GoMaxProcs: runtime.GOMAXPROCS(0),
	}

	// 获取内存信息
	var mem runtime.MemStats
	runtime.ReadMemStats(&mem)
	info.TotalMemory = mem.Sys
	info.FreeMemory = mem.Frees

	// 获取主机名
	if hostname, err := os.Hostname(); err == nil {
		info.Hostname = hostname
	}

	// 获取当前目录
	if cwd, err := os.Getwd(); err == nil {
		info.CurrentDir = cwd
	}

	// 获取临时目录
	info.TempDir = os.TempDir()

	return info
}

// collectGoInfo 收集Go信息
func collectGoInfo() GoInfo {
	return GoInfo{
		Version:  runtime.Version(),
		GOOS:     runtime.GOOS,
		GOARCH:   runtime.GOARCH,
		Compiler: runtime.Compiler,
	}
}

// collectNodeInfo 收集Node.js信息
func collectNodeInfo() NodeInfo {
	info := NodeInfo{Installed: false}

	// 查找node
	path, err := exec.LookPath("node")
	if err != nil {
		return info
	}

	info.Installed = true
	info.Path = path

	// 获取版本
	if version, err := exec.Command(path, "--version").Output(); err == nil {
		info.Version = strings.TrimSpace(string(version))
	}

	// 获取npm版本
	if npmPath, err := exec.LookPath("npm"); err == nil {
		if version, err := exec.Command(npmPath, "--version").Output(); err == nil {
			info.NpmVersion = strings.TrimSpace(string(version))
		}
	}

	return info
}

// collectEnvironment 收集环境变量
func collectEnvironment() map[string]string {
	envVars := []string{
		"PATH",
		"GOPATH",
		"GOROOT",
		"HOME",
		"USER",
		"TEMP",
		"TMP",
		"OFFICE_DEBUG",
		"NODE_ENV",
		"DEBUG",
	}

	env := make(map[string]string)
	for _, key := range envVars {
		if value := os.Getenv(key); value != "" {
			env[key] = value
		}
	}

	return env
}

// collectDependencies 收集依赖信息
func collectDependencies() []DependencyInfo {
	deps := []DependencyInfo{
		{Name: "excelize", Version: getExcelizeVersion(), Status: "ok"},
	}

	// 检查Mammoth
	mammothInfo := checkMammoth()
	deps = append(deps, mammothInfo)

	// 检查PDF.js
	pdfInfo := checkPDFJS()
	deps = append(deps, pdfInfo)

	return deps
}

// getExcelizeVersion 获取excelize版本
func getExcelizeVersion() string {
	return "v2.x (latest)"
}

// checkMammoth 检查Mammoth
func checkMammoth() DependencyInfo {
	info := DependencyInfo{
		Name:   "mammoth",
		Status: "not_found",
	}

	// 检查脚本是否存在
	scriptPath := "scripts/mammoth/package.json"
	if data, err := os.ReadFile(scriptPath); err == nil {
		var pkg struct {
			Version string `json:"version"`
		}
		if json.Unmarshal(data, &pkg) == nil {
			info.Version = pkg.Version
			info.Status = "ok"
			info.Location = scriptPath
		}
	}

	return info
}

// checkPDFJS 检查PDF.js
func checkPDFJS() DependencyInfo {
	info := DependencyInfo{
		Name:   "pdf.js",
		Status: "not_found",
	}

	// 检查脚本是否存在
	scriptPath := "scripts/pdfjs/package.json"
	if data, err := os.ReadFile(scriptPath); err == nil {
		var pkg struct {
			Dependencies map[string]string `json:"dependencies"`
		}
		if json.Unmarshal(data, &pkg) == nil {
			if version, ok := pkg.Dependencies["pdfjs-dist"]; ok {
				info.Version = version
				info.Status = "ok"
				info.Location = scriptPath
			}
		}
	}

	return info
}

// PrintDiagnostics 打印诊断信息
func PrintDiagnostics() {
	info := CollectDiagnostics()

	fmt.Println("=== Office文档库诊断信息 ===\n")

	// 系统信息
	fmt.Println("【系统信息】")
	fmt.Printf("  操作系统: %s/%s\n", info.SystemInfo.OS, info.SystemInfo.Arch)
	fmt.Printf("  CPU核心: %d\n", info.SystemInfo.NumCPU)
	fmt.Printf("  内存: %d MB\n", info.SystemInfo.TotalMemory/1024/1024)
	fmt.Printf("  主机名: %s\n", info.SystemInfo.Hostname)
	fmt.Printf("  当前目录: %s\n", info.SystemInfo.CurrentDir)
	fmt.Printf("  临时目录: %s\n", info.SystemInfo.TempDir)
	fmt.Println()

	// Go信息
	fmt.Println("【Go信息】")
	fmt.Printf("  版本: %s\n", info.GoInfo.Version)
	fmt.Printf("  平台: %s/%s\n", info.GoInfo.GOOS, info.GoInfo.GOARCH)
	fmt.Println()

	// Node.js信息
	fmt.Println("【Node.js信息】")
	if info.NodeInfo.Installed {
		fmt.Printf("  版本: %s\n", info.NodeInfo.Version)
		fmt.Printf("  路径: %s\n", info.NodeInfo.Path)
		if info.NodeInfo.NpmVersion != "" {
			fmt.Printf("  NPM版本: %s\n", info.NodeInfo.NpmVersion)
		}
	} else {
		fmt.Println("  未安装")
	}
	fmt.Println()

	// 依赖信息
	fmt.Println("【依赖检查】")
	for _, dep := range info.Dependencies {
		statusIcon := "✓"
		if dep.Status != "ok" {
			statusIcon = "✗"
		}
		fmt.Printf("  %s %s: %s", statusIcon, dep.Name, dep.Version)
		if dep.Location != "" {
			fmt.Printf(" (%s)", dep.Location)
		}
		fmt.Println()
	}
	fmt.Println()

	// 性能统计
	if len(info.PerfStats) > 0 {
		fmt.Println("【性能统计】")
		for _, stat := range info.PerfStats {
			fmt.Printf("  %s:\n", stat.Name)
			fmt.Printf("    调用次数: %d\n", stat.Count)
			fmt.Printf("    总耗时: %s\n", stat.Total)
			fmt.Printf("    平均耗时: %s\n", stat.Avg)
			fmt.Printf("    最小耗时: %s\n", stat.Min)
			fmt.Printf("    最大耗时: %s\n", stat.Max)
		}
		fmt.Println()
	}

	// 追踪统计
	if info.TracerStats.TotalSpans > 0 {
		fmt.Println("【追踪统计】")
		fmt.Printf("  总跨度数: %d\n", info.TracerStats.TotalSpans)
		fmt.Printf("  活跃跨度: %d\n", info.TracerStats.ActiveSpans)
		fmt.Printf("  成功率: %.2f%%\n", float64(info.TracerStats.SuccessCount)/float64(info.TracerStats.TotalSpans)*100)
		fmt.Printf("  总耗时: %s\n", info.TracerStats.TotalDuration)
		fmt.Println()
	}

	// 最近错误
	if len(info.RecentErrors) > 0 {
		fmt.Println("【最近错误】")
		for i, entry := range info.RecentErrors {
			if i >= 5 {
				break
			}
			fmt.Printf("  [%d] %s [%s] %s\n", i+1, entry.Time.Format("15:04:05"), entry.Category, entry.Message)
			if entry.Error != "" {
				fmt.Printf("      %s\n", entry.Error)
			}
		}
		fmt.Println()
	}
}

// ExportDiagnostics 导出诊断信息
func ExportDiagnostics(format string) ([]byte, error) {
	info := CollectDiagnostics()

	switch format {
	case "json":
		return json.MarshalIndent(info, "", "  ")
	case "text":
		return []byte(formatDiagnosticsText(info)), nil
	default:
		return json.MarshalIndent(info, "", "  ")
	}
}

// formatDiagnosticsText 格式化诊断信息为文本
func formatDiagnosticsText(info *DiagnosticInfo) string {
	var sb strings.Builder

	sb.WriteString("=== Office文档库诊断信息 ===\n\n")

	// 系统信息
	sb.WriteString("【系统信息】\n")
	sb.WriteString(fmt.Sprintf("  操作系统: %s/%s\n", info.SystemInfo.OS, info.SystemInfo.Arch))
	sb.WriteString(fmt.Sprintf("  CPU核心: %d\n", info.SystemInfo.NumCPU))
	sb.WriteString(fmt.Sprintf("  内存: %d MB\n", info.SystemInfo.TotalMemory/1024/1024))
	sb.WriteString(fmt.Sprintf("  主机名: %s\n", info.SystemInfo.Hostname))
	sb.WriteString("\n")

	// 依赖信息
	sb.WriteString("【依赖检查】\n")
	for _, dep := range info.Dependencies {
		status := "OK"
		if dep.Status != "ok" {
			status = "NOT FOUND"
		}
		sb.WriteString(fmt.Sprintf("  %s: %s %s\n", dep.Name, dep.Version, status))
	}
	sb.WriteString("\n")

	// 性能统计
	if len(info.PerfStats) > 0 {
		sb.WriteString("【性能统计】\n")
		for _, stat := range info.PerfStats {
			sb.WriteString(fmt.Sprintf("  %s:\n", stat.Name))
			sb.WriteString(fmt.Sprintf("    调用次数: %d\n", stat.Count))
			sb.WriteString(fmt.Sprintf("    平均耗时: %s\n", stat.Avg))
		}
	}

	return sb.String()
}

// DiagnoseProblem 诊断问题
func DiagnoseProblem(problem string) string {
	var sb strings.Builder
	sb.WriteString(fmt.Sprintf("诊断: %s\n\n", problem))

	switch {
	case strings.Contains(problem, "docx"):
		sb.WriteString(diagnoseDOCX())
	case strings.Contains(problem, "excel"):
		sb.WriteString(diagnoseExcel())
	case strings.Contains(problem, "pptx"):
		sb.WriteString(diagnosePPTX())
	case strings.Contains(problem, "pdf"):
		sb.WriteString(diagnosePDF())
	case strings.Contains(problem, "node"):
		sb.WriteString(diagnoseNode())
	default:
		sb.WriteString("无法识别问题类型，请提供更多细节。\n")
	}

	return sb.String()
}

// diagnoseDOCX 诊断DOCX问题
func diagnoseDOCX() string {
	var sb strings.Builder
	sb.WriteString("【DOCX问题诊断】\n\n")

	nodeInfo := collectNodeInfo()

	if !nodeInfo.Installed {
		sb.WriteString("✗ Node.js未安装\n")
		sb.WriteString("  解决方案: 安装Node.js https://nodejs.org/\n\n")
	}

	// 检查Mammoth
	scriptPath := "scripts/mammoth/docx2html.js"
	if _, err := os.Stat(scriptPath); os.IsNotExist(err) {
		sb.WriteString("✗ Mammoth脚本不存在\n")
		sb.WriteString("  解决方案: cd scripts/mammoth && npm install\n\n")
	} else {
		sb.WriteString("✓ Mammoth脚本存在\n\n")
	}

	return sb.String()
}

// diagnoseExcel 诊断Excel问题
func diagnoseExcel() string {
	return "【Excel问题诊断】\n\n" +
		"Excel使用excelize库处理，该库为Go原生实现，无需外部依赖。\n" +
		"如果遇到问题，请检查:\n" +
		"1. 文件是否为有效的Excel文件\n" +
		"2. 文件是否被其他程序占用\n\n"
}

// diagnosePPTX 诊断PPTX问题
func diagnosePPTX() string {
	var sb strings.Builder
	sb.WriteString("【PPTX问题诊断】\n\n")

	nodeInfo := collectNodeInfo()

	if !nodeInfo.Installed {
		sb.WriteString("✗ Node.js未安装\n")
		sb.WriteString("  解决方案: 安装Node.js https://nodejs.org/\n\n")
	}

	// 检查Mammoth
	scriptPath := "scripts/mammoth/pptx2html.js"
	if _, err := os.Stat(scriptPath); os.IsNotExist(err) {
		sb.WriteString("✗ Mammoth脚本不存在\n")
		sb.WriteString("  解决方案: cd scripts/mammoth && npm install\n\n")
	} else {
		sb.WriteString("✓ Mammoth脚本存在\n\n")
	}

	return sb.String()
}

// diagnosePDF 诊断PDF问题
func diagnosePDF() string {
	return "【PDF问题诊断】\n\n" +
		"PDF使用pdf.js处理，需要Node.js环境。\n" +
		"如果遇到问题，请检查:\n" +
		"1. Node.js是否正确安装\n" +
		"2. pdfjs-dist包是否安装\n\n"
}

// diagnoseNode 诊断Node问题
func diagnoseNode() string {
	var sb strings.Builder
	sb.WriteString("【Node.js问题诊断】\n\n")

	nodeInfo := collectNodeInfo()

	if nodeInfo.Installed {
		sb.WriteString(fmt.Sprintf("✓ Node.js已安装: %s\n", nodeInfo.Version))
		sb.WriteString(fmt.Sprintf("  路径: %s\n", nodeInfo.Path))
		if nodeInfo.NpmVersion != "" {
			sb.WriteString(fmt.Sprintf("  NPM版本: %s\n", nodeInfo.NpmVersion))
		}
	} else {
		sb.WriteString("✗ Node.js未安装\n")
		sb.WriteString("  解决方案: 安装Node.js https://nodejs.org/\n")
		sb.WriteString("  安装后运行: npm install -g npm@latest\n\n")
	}

	return sb.String()
}

// CheckEnvironment 检查环境
func CheckEnvironment() (bool, []string) {
	var errors []string

	// 检查Node.js
	nodeInfo := collectNodeInfo()
	if !nodeInfo.Installed {
		errors = append(errors, "Node.js未安装")
	}

	// 检查Mammoth脚本
	scriptPath := "scripts/mammoth/package.json"
	if _, err := os.Stat(scriptPath); os.IsNotExist(err) {
		errors = append(errors, "Mammoth依赖未安装，请运行: cd scripts/mammoth && npm install")
	}

	// 检查目录权限
	tempDir := os.TempDir()
	if _, err := os.Create(tempDir + "/office-test"); err != nil {
		errors = append(errors, fmt.Sprintf("临时目录无写入权限: %s", tempDir))
	}

	return len(errors) == 0, errors
}

// GetSystemInfo 获取系统信息
func GetSystemInfo() SystemInfo {
	return collectSystemInfo()
}

// GetNodeInfo 获取Node信息
func GetNodeInfo() NodeInfo {
	return collectNodeInfo()
}

// 辅助函数

// getDependencyVersions 获取依赖版本
func getDependencyVersions() map[string]string {
	versions := make(map[string]string)

	// excelize
	versions["excelize"] = getExcelizeVersion()

	// mammoth
	mammothInfo := checkMammoth()
	if mammothInfo.Status == "ok" {
		versions["mammoth"] = mammothInfo.Version
	}

	// pdf.js
	pdfInfo := checkPDFJS()
	if pdfInfo.Status == "ok" {
		versions["pdfjs-dist"] = pdfInfo.Version
	}

	return versions
}

// GetDependencyInfo 获取依赖信息
func GetDependencyInfo() []DependencyInfo {
	return collectDependencies()
}

// VerifyInstallation 验证安装
func VerifyInstallation() (bool, error) {
	// 检查环境
	ok, errors := CheckEnvironment()
	if !ok {
		return false, fmt.Errorf("环境检查失败: %s", strings.Join(errors, "; "))
	}

	// 检查Node.js
	nodeInfo := collectNodeInfo()
	if !nodeInfo.Installed {
		return false, fmt.Errorf("Node.js未安装")
	}

	return true, nil
}

// PrintVersionInfo 打印版本信息
func PrintVersionInfo() {
	fmt.Println("=== Office文档库版本信息 ===")
	fmt.Println()

	versions := getDependencyVersions()

	fmt.Println("【依赖版本】")
	if len(versions) > 0 {
		keys := make([]string, 0, len(versions))
		for k := range versions {
			keys = append(keys, k)
		}
		sort.Strings(keys)

		for _, k := range keys {
			fmt.Printf("  %s: %s\n", k, versions[k])
		}
	} else {
		fmt.Println("  无依赖信息")
	}
	fmt.Println()

	// Go版本
	fmt.Println("【运行时】")
	fmt.Printf("  Go: %s\n", runtime.Version())
	nodeInfo := collectNodeInfo()
	if nodeInfo.Installed {
		fmt.Printf("  Node.js: %s\n", nodeInfo.Version)
		if nodeInfo.NpmVersion != "" {
			fmt.Printf("  NPM: %s\n", nodeInfo.NpmVersion)
		}
	} else {
		fmt.Println("  Node.js: 未安装")
	}
}
