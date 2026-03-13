---
name: qa-defect-report
description: 读取本地 Excel 格式缺陷清单，自动完成全套缺陷分析并生成结构化报告（Markdown + HTML + PDF）。分析维度包括：缺陷概览统计、各版本/模块分布、9类根因自动归类、版本质量趋势评级、AI修复与测试改进建议。当用户提供缺陷 Excel 文件路径并要求做缺陷分析、版本质量分析、出缺陷报告时触发。
---

# QA 缺陷分析报告 (qa-defect-report)

从本地 Excel 缺陷文件一键生成完整缺陷分析报告，输出 Markdown + HTML + PDF 三格式。

## 执行流程

### Step 0 — 确认输入文件路径（必须）

**文件路径由用户在对话中指定，不写死。** 获取顺序：

1. **对话中已给路径**：直接使用（如"分析 `D:\项目\缺陷.xlsx`"）
2. **IDE 中当前打开的 xlsx 文件**：从上下文感知，询问用户确认
3. **未提供路径**：直接问用户：

   > "请告诉我缺陷清单 Excel 文件的完整路径（如 `D:\测试\缺陷.xlsx`），以及报告需要使用的项目名称（用于报告标题和文件命名）。"

确认两个变量后再继续：
- `$xlsxPath`：Excel 文件完整路径，如 `D:\测试缺陷分析\晋建保理QA环境缺陷.xlsx`
- `$projectName`：项目名，如 `晋建保理`，用于报告标题和输出文件名
- `$outputDir`：默认 = Excel 文件所在目录（`Split-Path $xlsxPath`），用户可覆盖

### Step 1 — 探测环境

```powershell
# 1. 找 Python（固定路径优先）
$py = "C:\Users\$env:USERNAME\AppData\Local\Programs\Python\Python312\python.exe"
if (!(Test-Path $py)) { $py = (Get-Command python -ErrorAction SilentlyContinue)?.Source }

# 2. 找 Pandoc
$pandoc = "C:\Program Files\Pandoc\pandoc.exe"
```

### Step 2 — 解析 xlsx

xlsx 本质是 ZIP+XML，**无需 pandas**，直接用 PowerShell 解析 `xl/worksheets/sheet1.xml`：

```powershell
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip  = [System.IO.Compression.ZipFile]::OpenRead($xlsxPath)
$entry = $zip.GetEntry("xl/worksheets/sheet1.xml")
$ms   = New-Object System.IO.MemoryStream
$entry.Open().CopyTo($ms)
$xml  = [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
$zip.Dispose()

# 每个单元格取内联字符串 <is><t>text</t></is>
# 列号 = cell ref 去掉数字，如 "AB2" → "AB"
$col = $ref -replace '\d+', ''
```

保存为 `<输出目录>/defects_raw.json`（utf-8-sig）供 Python 读取。

### Step 3 — 字段映射（标准列定义）

| 列号 | 字段 | 含义 |
|------|------|------|
| A | ID | 缺陷编号 |
| B | Title | 缺陷标题 |
| C | StoryName | 关联Story |
| F | Severity | 严重程度（致命/严重/一般/轻微） |
| G | Priority | 优先级 |
| H | DiscoverPhase | 发现阶段 |
| I | InjectPhase | 注入阶段 |
| L | DiscoverVer | 发现版本 |
| M | FixVer | 修复版本 |
| S | Desc | 缺陷描述 |
| T | DefectType | 缺陷类型 |
| U | Status | 缺陷状态（关闭/已拒绝/待修复） |
| X | Module | 影响模块 |
| Y | RootCause | 原因及分析 |

> 如果目标文件列顺序不同，在执行前先确认列号映射。

### Step 4 — Python 分析脚本

用 `Write` 工具将以下脚本写到 `%TEMP%\defect_full_analysis.py`，再执行：

**核心逻辑（必须包含）：**

```python
# 1. 读 JSON（utf-8-sig 兼容 BOM）
with open(json_file, "r", encoding="utf-8-sig") as f:
    defects = json.load(f)

# 2. 严重级别标准化 → 致命 / 严重 / 一般 / 轻微/提示
SEVERITY_MAP 关键字匹配：
  "致命" / "fatal" / "blocker"  → 致命
  "严重" / "critical" / "major" → 严重
  "轻微" / "minor" / "小"       → 轻微/提示
  其余                           → 一般

# 3. 版本号排序
def parse_ver(v): 提取 (major, minor, patch) 三元组，用于时序排列

# 4. 质量评级（每个版本）
def quality_grade(total, high_sev):
    high_ratio = high_sev / total
    if total > 100 or high_ratio > 0.25: return "差"
    elif total > 60 or high_ratio > 0.15: return "一般"
    else: return "良好"
```

**根因 9 分类（关键词正则，按优先级匹配）：**

| 优先级 | 分类 | 匹配关键词（含中英文） |
|--------|------|----------------------|
| 1 | 数据库/数据同步问题 | `sql\|SQL\|数据库\|数据同步\|脚本未同步\|Redis\|ZK\|Zookeeper\|数据异常` |
| 2 | 接口/参数问题 | `接口\|参数\|api\|API\|入参\|出参\|字段\|协议\|路由\|接口.*改动` |
| 3 | 边界/空值未处理 | `为空\|空值\|Null\|NULL\|边界\|越界\|缺失\|未判断\|没有判断\|没有处理` |
| 4 | 前端/UI 问题 | `前端\|页面\|UI\|loading\|样式\|显示\|展示\|渲染\|兼容\|Front` |
| 5 | 配置/部署问题 | `配置\|部署\|Config\|脚本\|未设置\|初始化\|环境\|未配置` |
| 6 | 权限/鉴权问题 | `权限\|鉴权\|认证\|授权\|CFCA\|证书\|角色.*权` |
| 7 | 性能/并发问题 | `性能\|并发\|超时\|timeout\|内存\|OOM\|死锁\|Perf` |
| 8 | 需求理解不一致 | `需求\|理解\|业务规则\|逻辑不清\|理解不一致\|规则.*遗漏` |
| 9 | 代码逻辑错误 | 有内容但未命中以上规则（兜底分类） |

### Step 5 — 报告五章结构

```
一、缺陷概览
  1. 总缺陷数（总量/关闭率/拒绝率）
  2. 各版本缺陷分布 TOP20（含高危占比、质量评级）
  3. 各模块/服务缺陷分布（按 Module 字段 + 微服务 regex 归类）
  4. 严重级别分布 + 缺陷类型分布

二、缺陷根因分析（自动归类）
  - 9类根因统计表（数量+占比）
  - 注入阶段分布
  - 每类：数量、典型案例（≤3条）、根本原因、开发改进、测试改进

三、版本质量分析
  1. 版本时序趋势表（总缺陷/高危数/高危占比/环比趋势/质量评级）
  2. 高风险模块（服务级，缺陷数 ≥ 10）
  3. 高频问题类型 TOP5
  4. 整体质量结论（差/一般/良好 + 支撑数据）

四、AI修复建议 & 测试改进建议
  - P0/P1/P2 修复优先级表
  - 8大测试加强场景
  - 需求/设计/编码/用例/回归 五阶段流程改进

五、可直接汇报的总结
  - 一句话质量结论
  - 3条最重要改进措施
```

### Step 6 — 输出文件

报告文件统一输出到 **Excel 文件所在目录**，命名格式：`{项目名}测试缺陷分析报告.{ext}`

| 格式 | 生成方式 |
|------|----------|
| `.md` | Python 直接写入（utf-8） |
| `.html` | `pandoc <md> -t html5 --standalone` + 注入中文样式 CSS |
| `.pdf` | 优先 Edge 无头模式：`msedge.exe --headless --print-to-pdf=<path> file:///<html>` |

**PDF 降级方案（按顺序尝试）：**
1. weasyprint（`pip install weasyprint`，需 GTK 依赖）
2. Edge 无头：`C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`
3. Chrome 无头：`C:\Program Files\Google\Chrome\Application\chrome.exe`
4. pandoc + xelatex（需 TeX Live）

### Step 7 — 中文编码注意事项

- Python 脚本用 `Write` 工具写入（正确 UTF-8），**不用 Shell heredoc**（会乱码）
- Python 脚本内 `print()` 避免非 ASCII 字符，或在文件开头 `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')`
- PowerShell 写文件：`[System.IO.File]::WriteAllText(path, content, [System.Text.Encoding]::UTF8)`
- Python 读 JSON：`encoding="utf-8-sig"`（兼容 PowerShell 写入的 BOM）

## 适配其他项目

当缺陷文件来自不同系统时，先确认：

1. **列号映射**：打印 header row，对照上面的字段映射表
2. **严重级别值**：确认实际值（如"Critical/High/Medium/Low" vs "致命/严重/一般/轻微"）
3. **版本字段**：确认版本号格式（如 `V6.3.0` vs `release-2024-03`）
4. **服务模块**：根据实际微服务名称调整 service regex 匹配规则

## 快速触发命令

用户说以下任何一句时自动应用本 skill：

```
# 带路径（推荐，直接开始不追问）
分析 D:\测试\项目缺陷.xlsx，出缺陷分析报告
读取 D:\xxx\缺陷清单.xlsx，做版本质量分析
帮我分析这个缺陷文件：D:\...

# 不带路径（Agent 会追问）
帮我做缺陷分析
读取缺陷清单，出分析报告
对这个 Excel 做版本质量分析
生成缺陷分析报告
```

> **路径规则**：文件路径在对话中给定即可，无需放到任何固定目录。报告默认输出到 Excel 文件所在目录。
