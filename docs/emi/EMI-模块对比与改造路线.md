# EMI 报告模块：EMID 与 ORTE 的能力差距、改造成本与路线

> **这份文档回答什么**：ORTE 的 EMI 模块到底差在哪、值不值得改、把 EMID 做成依赖库划不划算、ORTE 的数据是怎么流动的、不同入口会带来什么差异、现成的工具有哪些能用。
> **范围**：分析（as-is + 选项评估），不含代码改动。
> **依据**：ORTE 仓库 `master@cec1624`（程序集 0.6.2.1）、EMID `D:\Desktop_Li\code\EMI_deal` v2.9（2026-10-17）。证据以 `文件:行号` 标注。
> **术语**：**EMID** = EMI_deal 制作程序（Python/PySide6）；**ORTE** = ORT 一键报告的 EMI 制作模块（.NET Framework 4.8 / WPF）。

---

## 0. 结论先看

1. **ORTE 的 EMI 不是"细节没完善"，而是"领域层没建起来"**。它是 EMID 早期版本的机械移植，移植时丢了模板族、写前校验、写后诊断、描述审计、图片匹配、状态记忆、批处理这七项；更硬的是三处契约断裂：随包模板没有 `Setup` 表（点「制作」只会弹"模板格式不符"）、Setup 里的行列坐标即使配上也不会生效、按文件名收集序列号/电压/负载的 `else if` 链会漏收（§3）。
2. **有改进的必要，但动作应该是"在 ORTE 内重建 EMI 领域层"，不是"修补这个 ViewModel"**。UI 外壳、NPOI 写入、OLE 嵌入、Excel/Word COM 这些资产都在，缺的是"输入装配 → 校验 → 写表 → 自检"这条可复用的服务链路。
3. **把 EMID 做成依赖库，成本高于收益**。EMID 有 56% 的代码与 Qt 强绑（连批处理都得造一个真实 QWidget 并给 `QMessageBox` 打补丁才能复用），再加上两套运行时、两条构建链、两个模板目录，以及它与 ORTE"从数据库拿权威值"的根本冲突。**不建议库化/嵌入式集成**。
4. 如果短期内要"两边都省事"，最划算的中间态是**轻集成**：给 EMID 补一个可发布的 `--job` CLI 入口，ORTE 只负责"装配作业 JSON → 调用 → 结果登记"，零重写就能拿到 EMID 的全部制作能力（前提是接受两个程序、两套模板目录）。

---

## 1. 两个东西是什么

| | **EMID**（EMI_deal） | **ORTE 的 EMI 模块** |
|---|---|---|
| 定位 | 单用途桌面工具：EMI 报告全流程 | ORT 管理系统里四类报告之一（Thermal Shock / Burn In / EMI / ATE） |
| 输入 | EMI 数据文件夹（`序列号-电压-负载-线性.pdf`）+ 模板 + 界面手填（工令/版本/周期/日期/描述/图片） | 同左，但工令/版本/DC 可由报告概览或领用表带入；模板来自 `Templates` |
| 输出 | `2.1 Conducted EMI Measurement_{型号}.xlsx` + `{型号}.zip` 并 OLE 嵌入 | 同左（`EMIReportViewModel.cs:577-591`） |
| 技术栈 | Python 3.8 · PySide6 · openpyxl · pdfplumber · Pillow · pywin32 | .NET Framework 4.8 · WPF · NPOI · DocumentFormat.OpenXml · iText 7 · Excel/Word COM |
| 规模 | 手写 ≈5,200 行（`src/gui/res_rc.py` 30,235 行是 qrc 产物，不计） | EMI 专属 ≈2,000 有效行（VM 776 / 工具 643 / 界面 509 / 模型 69 / 服务 17）+ 可复用共用基础设施 ≈8,900 有效行 |
| 成熟度 | 18 项成品级能力，含无界面批处理、借用数据、替换序列号、数值抖动 | 约 EMID 2.0 时代的功能面（写一张表），且契约断裂 |

---

## 2. 能力差距矩阵

图例：**●** 完整可用 · **◐** 部分/需手工 · **○** 缺失

| # | 能力 | EMID | ORTE | ORTE 证据 / 缺口 |
|---|---|:--:|:--:|---|
| 1 | PDF 直读测量数据 | ● | ○ | ORTE 只读 docx 表格（`SmartTableExtractor.ConvertWordTablesToCsv`），PDF 仅用于打包；iText 7 已在引用但只被死代码 `Main/Views/WindowTestPdf.xaml.cs:27-37` 使用 |
| 2 | 模板族（负载档数 1/3/4/10 + AMZ + Vdc） | ● 6 套 | ○ 1 套 | 仓库内 `Templates/# ORT Test Report (WK#)_#/Report/2.1 Conducted EMI Measurement.xlsx` **只有 `Conducted EMI` 一张表、没有 `Setup`**，而代码要求两张表 |
| 3 | 模板自描述锚点（Setup 表驱动） | ● | ○ | ORTE 把它改成"B1 存 JSON"（`EMISetupViewModel.cs:36-52`），与 EMID 模板的 `B1=44`（SN 起始行）冲突；且读/写键名不一致 + 嵌套类型判定错误，**配置从不生效**（§3.4） |
| 4 | 写前条件矩阵硬校验 | ● | ○ | EMID `EMI-Report.py:1566-1645`；ORTE 直接写，条件对不上就写歪 |
| 5 | 写后完整性检查 + 缺失归因 | ● | ○ | EMID `EMI-Report.py:1768-1854`；ORTE 无 |
| 6 | 保格式删行（合并区/行高/公式/边框） | ● | ● | `Utils/ExcelNpoi.cs:1122-1192` 已实现同类能力（合并区重算 + 移位） |
| 7 | 测试描述审计（电压/负载按数据改写） | ● | ○ | EMID `src/desc_audit.py`（345 行纯函数）；ORTE 描述原样写 `F8` |
| 8 | 测试图片自动匹配（机种/系列/日期） | ● | ○ | EMID `src/picmatch.py`；ORTE 的 EMI 页不从图片目录取图 |
| 9 | 图片等比缩放插入（高 250 / 锚点 C27） | ● | ○ | ORTE 的图片能力在 Thermal/BurnIn 路径（`ExcelAddPicture`），EMI 路径未接 |
| 10 | 界面状态记忆（崩溃可恢复） | ● | ◐ | ORTE 只持久化 `EmiDataPath`（`Models/LocalSettings.cs:29`），其余每次重填 |
| 11 | 序列表格 + 按序列号覆盖工令/版本 | ● | ○ | ORTE 每个 SN 块用工令/版本/DC 的链式公式引用首行（`EMIReportViewModel.cs:345-350`），无覆盖入口 |
| 12 | 加载 CSV/批量 | ● | ○ | ORTE 无批处理入口；`ReportGenerationService.GenerateEMIReport` 只 `throw NotImplementedException`（`Services/ReportGenerationService.cs:229-238`） |
| 13 | 数据来源权威性 | ○ 文件名/PDF 推断 | ● 数据库 + 领用表 | ORTE 可由 `plan`/`requisition` 带入工令/版本/DC/周期（`Plans/Views/WindowPlans.xaml.cs:2136-2148`） |
| 14 | 与系统其它模块联动 | ○ | ● | 计划右键进入、报告扫描回写状态、权限控制、托盘入口 |
| 15 | OLE 附件嵌入 | ● | ● | EMID 用 pywin32；ORTE 用 `Utils/ExcelOleEmbedder`（Excel COM，`ExcelNpoi` 无 OLE 写能力） |
| 16 | docx↔pdf、Word 文本替换 | ● | ● | EMID `docx2pdf.py` + `gui/ui_wordreplace.py`；ORTE `Docx2Pdf` + `DocxTextReplacer` + `WindowEmiTools` |
| 17 | 报告扫描判定完成（TestStatus） | ○ | ● | ORTE `Services/ReportStatusReader.cs`（`UNITS UNDER TEST == TOTAL + FAIL`） |
| 18 | 报告骨架/以已有报告为模板 | ○ | ● | ORTE `Services/ReportTemplateService.cs`（1,921 行，含排期与整目录复制） |

**一句话**：ORTE 强在"系统集成与数据权威性"，EMID 强在"制作链路的细节与自检"。当前 ORTE 的 EMI 恰好把两者的强项都丢了——既没接住数据库的优势场景，也没复刻 EMID 的制作细节。

---

## 3. ORTE 现状：哪里断了（可复现证据）

### 3.1 模板契约断裂（最硬的一条）

- 代码要求 `Conducted EMI` + `Setup` 两张表，缺一张就弹错返回：`EMIReportViewModel.cs:557-567`。
- 仓库内随程序发布的 EMI 模板只有 `Conducted EMI` 一张表（实测 `Templates/# ORT Test Report (WK#)_#/Report/2.1 Conducted EMI Measurement.xlsx` 的 `xl/workbook.xml` 只有一个 `<sheet>`，`A1:AA103`）。
- 想让它跑通，只能**先用「EMI 设置」对话框把 JSON 另存成一个新 xlsx**（`EMISetupViewModel.cs:54-73` 走 `SaveJsonToExcel`，会 `CreateSheet("Setup")`），再用「选择模板」选中那个文件。也就是说：默认路径点「制作」必然失败，能成功的那条路要先绕一次"设置"。
- 若直接换成 EMID 模板（`Setup!B1=44` 而不是 JSON），`SettingsViewModel.ParseJson` 用 Newtonsoft `JObject.Parse`（`Main/ViewModels/SettingsViewModel.cs:63-76`）会抛异常，而 `DoReport` 这段只有 `try/finally` 没有 `catch`（`EMIReportViewModel.cs:552-583`）→ 异常冒泡到 `async void`。
- 结论：**默认模板不可用；即使绕出 Setup 表，坐标映射也不生效（见 3.4）**，实际仍走硬编码坐标。

### 3.2 模板自动选择会选错

`GetEMITemplatePath` 在 `TemplateDir` 下递归取**所有** xlsx，用"文件名包含 `-{电压}_{负载}`"匹配，匹配不到就**回退 `excelFiles[0]`**（`EMIReportViewModel.cs:220-233`）。`Templates` 根下还有 `1.1 Thermal Shock Test.xlsx`、`1.2 ORT Burn-In Test Report.xlsx`、`ATE report.xlsx`，回退很可能选到非 EMI 模板。

### 3.3 序列号/电压/负载收集有逻辑缺陷

```csharp
if (!emiUUTdatasInfo.SN.Contains(infos[0])) emiUUTdatasInfo.SN.Add(infos[0]);
else if (!emiUUTdatasInfo.Voltage.Contains(infos[1])) emiUUTdatasInfo.Voltage.Add(infos[1]);
else if (!emiUUTdatasInfo.Load.Contains(infos[2])) emiUUTdatasInfo.Load.Add(infos[2]);
```
（`EMIReportViewModel.cs:134-137`）

`else if` 链意味着：只有"SN 已存在"时才去收集电压，只有"电压已存在"时才收集负载——首份文件只登记 SN，收集结果**取决于文件枚举顺序**；并且 `infos[0..3]` 没有长度检查，文件名不含 3 个 `-` 会 `IndexOutOfRange`。同处的相线判定 `line.Contains("L1")` / `line.Contains("N")`（`:175-176`）也会被任意含 `N` 的文字命中。

### 3.4 Setup 配置链路三处一起失效

| 失效点 | 证据 | 后果 |
|---|---|---|
| 读侧键名是下划线、写侧键名是空格 | 读：`EMIReportViewModel.cs:282-287`（`TESTED_BY`…）；写：`EMISetupViewModel.cs:20`（`"TESTED BY"`…） | 第一个取值即 `KeyNotFoundException`，被 `:290` 的 `catch` 吞掉，静默退回硬编码地址 |
| 嵌套对象类型不匹配 | `SettingsViewModel.cs:71-73` 把子对象转成 `Dictionary<string,object>`，再下一层 `Row`/`Col` 仍是 `JObject`；而 `EMIReportViewModel.cs:261` 用 `is Dictionary<string, object>` 判定 | 行列映射**从不生效**，无论用户怎么在设置窗口改 |
| 默认 JSON 本身不合法 | `EMISetupViewModel.cs:20`：`"TEST STAGE"` 重复、结尾 `"TEST CONCLUSION":"M6",}` 多逗号 | `LoadFromJson` 抛错被 `catch`（`SettingsViewModel.cs:86-94`），表现为"默认设置加载失败、界面空表" |

结论：这套"模板自描述锚点"在 ORTE 里是**装饰性的**——界面能改、文件能存，生成时不读。

### 3.5 两套并行的生成实现，服务版整条是死路

- `ReportGenerationService` 的注释直接写明"待迁移：`EMIReportViewModel.DoReport` → `GenerateEMIReport`"，方法体 `throw new NotImplementedException`（`Services/ReportGenerationService.cs:229-238`）。
- 更关键的是：**这个服务在 DI 里注册了（`App.xaml.cs:124`），却没有任何调用方**。Thermal/BurnIn 的真实生成路径是 `BaseReportPageViewModel.ReportFinish`（`Reports/ViewModels/BaseReportPageViewModel.cs:55-190`），与 `ReportGenerationService.GenerateBaseReport`（`:63-185`）是**逐行同构的两份实现**。
- 后果：
  - EMI 是四类报告里唯一"没有服务层实现"的（`ReportGenerationService.cs:222-238` 只有 17 行占位），能力全压在 ViewModel（`WriteDatas` 一个方法 277 行）；
  - 计划右键带入的 `PrefilledReportModel` 永远是 `BurnInReportModel`（`WindowPlans.xaml.cs:2172-2187`），EMI 拿不到；
  - 未来要做"批量/无人值守"，EMI 无处挂；而就算想接，也得先把这套"注册了没人用"的服务复活。

### 3.6 生成完不回写，闭环是断的

- EMI 保存后只弹一个 `MessageBox`：不写 `report_links`、不改 `plans.ReportStatus`、不校验输出目录（`EMIReportViewModel.cs:577-591`）。
- 报告状态只由扫描回写，且判定依据是**概览的 `TestStatus` 表**，与 EMI 工作簿是否存在无关（`Services/ReportStatusReader.cs:57-108`，完成条件 `UNITS UNDER TEST == TOTAL + FAIL`）。
- 因此"生成 EMI 报告"与"系统认为 EMI 已完成"是两条互不相干的链路。

### 3.7 其它历史包袱

| 痕迹 | 位置 | 影响 |
|---|---|---|
| 死窗口 + 唯一的 iText 用法 | `Main/Views/WindowTestPdf.xaml.cs`（全仓无引用） | iText 7 被打进发布包，只为一段试验代码 |
| 死字段：EMI 页的 WorkOrder/Version/DC 输入框 | `UserControlEMIReportPage.xaml.cs:142-144` 只写入 VM 属性；生成时读的是 `UUTInfos`（`EMIReportViewModel.cs:302-304`） | 用户手改无效，且看不出无效 |
| `emiVM` 静态单例 + DI `AddTransient` | `UserControlEMIReportPage.xaml.cs:26,55`、`App.xaml.cs:140` | 重建 Tab / 多窗口时状态串味 |
| 写了一半的死代码 | `EMIReportViewModel.cs:306-324`（`sn_rows`/`SN_cells` 构造后不用）、`:30,128,142`（`emiPdfFiles` 只清不加） | 误导读代码的人 |
| 无效兜底路径 | `EMIReportViewModel.cs:577`：`Directory.GetCurrentDirectory() + "2.1 ...xlsx"` 少分隔符 | 兜底时会把文件拼到目录名后面 |
| 无调用方的方法 | `Utils/Docx2Pdf.cs:214-245`（`AlertFileTime`，且 `:228` 对 `DateTime` 的 `AddDays` 结果被丢弃） | 历史迁移残留 |
| 硬编码 UNC 路径的死字段 | `WindowMainReport.xaml.cs:34-40`（`\\bnt56\...`） | 无引用 |
| EMI 数据处理入口藏在"数据路径标签连击三下" | `UserControlEMIReportPage.xaml.cs:61-83`（原按钮 `Collapsed`） | 可发现性差 |
| EMI Tab 默认不创建 | `WindowMainReport.xaml.cs:52-57`（`DefaultChecked=false`） | 从计划右键进入时若没勾过"视图→EMI"，用户看不到 EMI 页 |
| EMI 页不参与单体数据填充 | `WindowMainReport.xaml.cs:327-336` | EMI 的 SN 与其它 Tab 的 SN 来自两套来源 |
| 魔数：频率、行高、字号、底色 | `EMIReportViewModel.cs:355`（电压含 "110" → 60Hz）、`:373`（21.75）、`:433-478` | 换模板/换标准要改代码 |

---

## 4. 有没有改进必要

**有，而且是刚需**。判断依据：

1. **正确性**：§3.1-3.3 是"会写错报告或直接在用户面前抛异常"级别的缺陷，不是体验问题。
2. **一致性**：四类报告里只有 EMI 没有服务层实现（`ReportGenerationService.cs:222-238` 是 17 行占位），其余三类也是"两份同构实现"（`BaseReportPageViewModel.ReportFinish` 与 `ReportGenerationService.GenerateBaseReport`，后者无人调用）。"模型化"这件事历史上只做了一半，EMI 是欠得最多的那个。每新增一个入口（托盘/计划右键/批处理/定时），EMI 都要在 ViewModel 里单独改一遍。
3. **成本不对称**：ORTE 已经付过的成本（NPOI 保格式删行、OLE 嵌入、Excel/Word COM、报告扫描、报告模板工具、权限、多语言）在 EMI 路径上**全部没用上**，等于白买。
4. **用户意图**：ORTE 的 EMI 是"辅助功能"，要的是"少填、少错、有提示"，而不是 100% 复刻 EMID。这恰好是"领域层 + 薄 UI"能满足、而"继续往 ViewModel 里堆逻辑"满足不了的。

**什么时候不值得改**（反向条件）：如果 EMI 报告每月只有个位数、且现场始终用 EMID 出报告、ORTE 只需要"登记与扫描"，那 ORTE 侧把 EMI Tab 隐藏、只保留报告扫描即可——但按现状（Tab 存在、能点、会抛错）反而更糟。

---

## 5. 「EMID 做成依赖库」成本评估

### 5.1 四种接法

| 方案 | 做法 | 边界 | 主要成本 | 主要风险 |
|---|---|---|---|---|
| **A. 进程外调用**（推荐度：中） | 调 EMID 的 CLI（`batch.py --job jobs.json`），ORTE `Process.Start` + JSON 交换 | 进程 | **EMID 当前没有可发布的 CLI**：`EMIRper.spec` 只打包 `src/EMI-Report.py`，`batch.py` 不进包（实测 `dist\EMIRper` 与 `00-EMIRper` 里都没有），要先在 EMID 侧补一个 `--job` 入口；之后是随包分发 ≈200MB 级运行时、冷启动 2–5s、无进度/取消协议 | 双运行时部署；EMID 升级要重新打包 |
| **B. 嵌入式 Python**（pythonnet / 嵌入解释器） | ORTE 进程内加载 Python | 进程内 | 版本矩阵（net48 × Python 3.8/3.11 × pythonnet）；Python DLL 与 stdlib 目录随包 | 崩溃即拖垮主程序；调试困难 |
| **C. 本地服务**（HTTP/命名管道常驻） | Python 起服务，ORTE 发请求 | 服务 | 端口/生命周期/开机自启/防火墙；异常恢复 | 多机部署每台都要起服务 |
| **D. 移植为 C# 领域层**（推荐度：高） | 参照 EMID 的算法在 ORTE 内重写（不是逐行翻译） | 无 | ≈1,500–2,500 行 C#；PDF 解析用已有 iText 7 | 两份实现可能漂移（可用"以 EMID 行为为验收基准"约束） |

### 5.2 为什么"库化"比看起来贵

- **代码边界不允许**：EMID 5,200 行里只有 22% 能零改动摘出；业务主干（`main_func` 及 14 个 helper，≈600 行）逻辑上不依赖 Qt，但形式上全是 `EMIWindow` 的方法、直接读控件、内联 `QMessageBox`。批处理为了复用，必须 `QT_QPA_PLATFORM=offscreen` 造真窗口再逐控件 `setText`、并替换 `QMessageBox`（`batch.py:58,115,522-541,704`）——这个"怪招"就是"业务与界面方法级粘连"的铁证。
- **契约要扩，不是复用**：EMID 从**文件名/PDF**推断序列号、机种、电压、负载；ORTE 的权威值在**数据库**（`plan`/`requisition`/报告概览）。要接上，就得给作业 JSON 增加"以 DB 为准"的覆盖字段与优先级规则——这部分是**新写**的，不是复用。
- **两份模板目录**：EMID 的 `template/`（6 套，`datas=[]` 不随包，必须放 exe 同级）与 ORTE 的 `Templates/`（报告骨架）语义不同。库化后要么合并（改 ORTE 的模板体系），要么双目录并存（用户要维护两处）。
- **演进速度不匹配**：EMID 在 2026-10-16/17 连续发了 2.3→2.9（批处理、借用数据、抖动、描述审计、另存数据）。库化会把"EMID 每次微调"变成"重新打包 + 更新 ORTE 发布包 + 回归"。

### 5.3 成本换算（相对值）

| 路线 | 一次性投入 | 持续成本 | 产出干净度 |
|---|---|---|---|
| A 进程外调用 | 中（打包/协议/错误映射/部署脚本） | 高（双栈、双构建） | 中（保留 EMID 行为，含其脆弱版式依赖） |
| D C# 领域层 | 中高（≈1,500–2,500 行 + 单测） | 低（单栈、可单测、跟 ORTE 一起发版） | 高 |
| 轻集成（见 5.4） | 低（EMID 补 CLI + ORTE 生成作业/调进程） | 中（两程序、两目录、两套 COM 依赖） | 中 |

### 5.4 唯一划算的"复用"姿势：轻集成（ORTE 调 EMID CLI）

不做库，做**进程 + 契约**。这条路成本远低于库化，前提是 EMID 侧补一个可发布的 CLI 入口（当前 `batch.py` 是开发态脚本，未进 `EMIRper.exe`）：

1. EMID 侧：给打包后的 `EMIRper.exe` 加 `--job jobs.json`（内部复用现有 `batch.py` 的作业解析与 `load_main_module`），stdout 输出结构化结果。
2. ORTE 侧：设置里记 EMID 路径 → 由计划/报告页生成 `jobs.json`（`data_dir`/`serials`/`template`/`load_qty`/`week`/`rev`/`workno`/`detail`/`test_date`/`image`/`out_dir`/`zip`/`embed`）→ `Process.Start` 等待 → 解析结果 → 用已有的报告扫描（`Services/ReportScanService.cs`）登记。

收益：**零重写**，立刻拥有 EMID 全部能力（模板族、条件校验、缺失诊断、描述审计、图片匹配）；EMID 升级不影响 ORTE。
代价：两个程序、两套模板目录、Excel COM 依赖两份；进度只能靠解析 stdout 或读 EMID 日志；用户机器上必须有 EMID。

**这条路的判断条件**：如果"EMID 长期仍是主工具、ORTE 的 EMI 只要够用"，选它；如果要"ORTE 独立成体系、可无人值守、可随 ORTE 一起发版"，选 5.1 的 D。

---

## 6. ORTE 的数据交互路径

### 6.1 四类数据源

| 源 | 内容 | 谁写入 | 位置 |
|---|---|---|---|
| SQLite 主库 | `plans`（机种/工令/阶段/起止/测试项）、`requisitions`（SN/工令/版本/DC）、`report_links`（报告文件夹） | 领退与计划模块 | `ort_plans.db` |
| 报告概览 xlsx | Cover（版本）、Waterfall（SN、工令、测试项与日期）、文件夹名里的 `WK####` | 测试员维护的概览文件 | `ReportDir` 下 |
| 本机设置 | `EmiDataPath`（EMI 数据目录）、`AteDataPath`、数据文件夹 | 设置窗口 | `Data/local_settings.json` |
| 模板 | `Templates/*.xlsx`（含 EMI 报告骨架与 Setup） | 随程序发布 / 报告模板工具生成 | 程序目录 |

### 6.2 装配链路（"生成报告"按钮）

```
WindowMainReport.DoReport_Click            Reports/Views/WindowMainReport.xaml.cs:338-394
  → HighCostTaskCoordinator.RequestStartAsync    （与扫描/索引互斥）
  → MainVM.ReadInfoFromOverview(ReportName)      Reports/ViewModels/MainReportViewModel.cs:144-332
      → 读 Cover/Waterfall → UUTInfoFromExcel{SNs, WorkOrder, Revision, DC, TestItems, TestPeriod, TestStart}
      → ApplyMatchedSourceToUUTInfos()           从计划进入时用领用表覆盖 SN/工令/版本/DC
  → 对每个已打开 Tab：ReadReportHeader() + SetReportResultData()
      → BaseReportPage.ReadReportHeader()        报告文件夹里的报告文件 → 回退模板
      → EMIReportPage.ReadReportHeader()         UserControlEMIReportPage.xaml.cs:87-169
```

**生成环节是另一条链**，与上面不共用服务：

| 报告类型 | 实际生成入口 | 服务层 |
|---|---|---|
| Thermal Shock / Burn In | `BaseReportPageViewModel.ReportFinish`（`Reports/ViewModels/BaseReportPageViewModel.cs:55-190`） | `ReportGenerationService.GenerateBaseReport`（`:63-185`）同构存在但**无调用方** |
| EMI | `EMIReportViewModel.DoReport`（`Reports/ViewModels/EMIReportViewModel.cs:541-592`） | `GenerateEMIReport` 只 `throw`（`:229-238`） |
| ATE | `WindowATE` 生成 ATE 文件，作为 OLE 附件交给 Thermal/BurnIn（`Reports/Views/WindowATE.xaml.cs`） | — |

### 6.3 EMI 自己的链路（与上面**并行**，不共享 SN）

```
选择数据路径（默认 EmiDataPath）             EMIReportViewModel.cs:607-611
  → ReadPath：从文件名拆 SN/电压/负载/相线   :121-146   ← 与 UUTInfos.SNs 是两套
选模板（可留空 → GetEMITemplatePath）
  → DoReport                                 :541-592
      → ReadDatas(docx)                       :148-218   SmartTableExtractor→CSV→Split(',')
      → 打开模板，读 Setup JSON               :568
      → WriteDatas：表头 + 分组写测量值 + 合并 + 样式 + 备注 + 删多余行  :235-511
      → 压缩 PDF 并生成 EMF 图标              :489-500   FileService.CreateFilteredZip
      → 保存 → ExcelOleEmbedder.Embed         :578-590
```

**关键问题**：表头来自"报告概览/领用表/报告文件夹"，数据来自"EMI 数据目录的文件名 + docx 内容"，中间没有任何一致性校验。两个来源打架时（例如概览里 SN 是 A/B，而数据目录里是 A/C），报告会同时包含两者且不报错。

### 6.4 回写

报告状态由**扫描**回写，不由生成环节回写：`Services/ReportScanService.cs` → 读概览 `TestStatus` 判定 → `plans.ReportStatus`（用户手工设"无要求"后锁定）。所以"生成完报告"与"报告状态变完成"是两条独立链路。

---

## 7. 不同入口的影响

| 入口 | 预填什么 | EMI Tab 是否创建 | 表头来源 | SN 来源 | 主要影响 / 风险 |
|---|---|---|---|---|---|
| 主窗口「一键报告」按钮 | 无（`ClearMatchedSource`，`MainWindow.xaml.cs:903`） | 否（默认不勾选，需"视图"菜单勾） | `RootPath` 下的报告模板 | 报告概览 | 直接进入时若上次是从计划进入，残留已被清理（好） |
| 托盘「一键报告 → EMI」 | 无 | **强制创建**并切页（`WindowMainReport.xaml.cs:199-249`） | 进页时**不读表头**（`SelectReportTab` 只在 `EnteredFromPlan` 时读），点「生成报告」后才读 | 同上 | 会顺手把"视图→EMI"勾选状态写回设置 |
| 计划表右键「打开一键报告」 | 计划 + 领用表 + 报告文件夹；`UUTInfos` 由领用表构造（`WindowPlans.xaml.cs:2081-2199`） | 否（仍看勾选） | **只从报告文件夹里的报告文件读**，读不到置空 | 领用表 SN（`ApplyMatchedSourceToUUTInfos`） | 若没勾过 EMI，进入后看不到 EMI 页；EMI 页的 SN 又不取领用表，而是重新从数据目录文件名推断 |
| 报告窗口内「选择报告概览」+「生成报告」 | 概览 | 同上 | 从计划进入且无报告文件 → 置空 | 概览（或领用表覆盖） | 会覆盖 `UUTInfos`，但不会动 EMI 自己的 `emiUUTdatasInfo` |
| EMI 页内「选择数据路径」 | 本机设置默认值 | 已存在 | — | **文件名** | 与上面所有 SN 来源并行、互不校验 |
| EMI 页「连击数据路径三次」 | 当前数据目录 | — | — | — | 只做两件事：改 PDF 文件时间、批量替换 docx 文本（`WindowEmiTools`），不参与生成 |
| 「报告模板工具」 | 机种/阶段/周期 | — | 生成整目录骨架（含 EMI 的 xlsx） | — | 生成的是**骨架**（无 Setup 表），不是可制作模板 |
| 报告扫描（后台） | — | — | — | — | 只回写状态，不参与制作 |

**结论**：入口差异主要落在三处——① 是否创建 EMI Tab；② 表头取"报告文件夹里的真实报告"还是"模板"；③ SN/工令/版本取"领用表"还是"概览"。而 EMI 的制作数据**不在这三条路径上**，它是第四条独立来源。这是当前设计最需要收口的地方。

---

## 8. 工具能力清单（现成可复用资产）

| 资产 | 能力 | EMI 可用性 |
|---|---|---|
| `Utils/ExcelNpoi.cs`（1,330 行） | NPOI 封装：读写、单元格、合并、样式 spec、行高、图片锚定、`InsertRows`/`DeleteRows`（含合并区重算，`:1122-1192`）、公式归一化 | **直接可用**，等价于 EMID 的 `delete_rows_keep_format` |
| `Utils/ExcelOleEmbedder.cs` + `ExcelOleWriter.cs` | Excel COM 嵌 OLE（多附件、图标、偏移、崩溃重启、Save 后时间戳校验）；**没有 Excel 时**由 `ExcelOleWriter` 直写 OpenXml 包兜底（手写 `xl/embeddings`、VML 图标、`oleObjects`） | 直接可用，且比 EMID 强（EMID 必须有 Excel） |
| `Utils/Docx2Pdf.cs` | Word COM/Interop 转 PDF；批量设文件时间 | 直接可用 |
| `Utils/DocxTextReplacer.cs` | docx 文本层替换（OpenXml） | 可直接扩展成"替换 + 抖动" |
| `Utils/SmartTableExtractor.cs` | docx 表格 → CSV（含嵌套） | 现被用于读测量表；注意 `Split(',')` 不处理引号转义 |
| `Utils/Report.cs` | 模板定位、表头读/写、周期解析、图片插入、地址换算 | 直接可用（EMI 已部分使用） |
| **iText 7**（已在 csproj） | PDF 页文本提取（`PdfTextExtractor.GetTextFromPage`） | **可直接替代 pdfplumber**，只需把版式常量搬过去 |
| `Utils/Image.cs` | 生成居中 EMF 图标、图片处理 | 直接可用 |
| `Services/FileService.cs` | 按正则过滤打包 zip | 直接可用 |
| `Services/ReportScanService.cs` | 扫描报告夹、判定完成、回写状态 | 直接可用（结果登记） |
| `Services/ReportTemplateService.cs`（1,921 行） | 生成机种报告骨架、以已有报告为模板、排期 | 生成的是骨架；EMI 的"制作模板"需另存 |
| `Services/HighCostTaskCoordinator.cs` | 高耗时任务互斥 | 新链路要接进来 |

**结论**：工具层几乎没有缺口，缺的是把它们编排成"EMI 制作服务"。

---

## 9. 建议路线

### 9.1 推荐：在 ORTE 内重建 EMI 领域层（方案 D）

**分层**（新建 `Services/Emi/`，ViewModel 退化成壳）：

```
EmiContext          输入装配：数据目录、序列号来源、机种、工令/版本/DC/周期/日期、描述、图片
EmiTemplateCatalog  模板选择：按负载档数/AMZ/Vdc 选定模板 + 锚点（读 Setup 表）
EmiMeasureParser    解析：PDF（iText）优先，docx 兜底 → EmiMeasurement[]
EmiReportWriter     写表：表头、条件矩阵校验、分组写入、合并、样式、备注、删行
EmiSelfCheck        自检：空条件删行、部分缺失诊断、描述审计、图片匹配
ReportGenerationService.GenerateEMIReport(model, path)   ← 接上模型化入口
```

**分阶段**（工作量为相对估算，不含联调与用户验收）：

| 阶段 | 内容 | 价值 | 规模 |
|---|---|---|---|
| **P0 打通** | 模板族与锚点契约统一；`Setup` 只作为锚点来源（不再塞 JSON，或升级 JSON 且兼容旧模板）；SN/电压/负载收集修正；模板选择改为"显式 + 校验"；补 `catch` 与友好提示 | 从"会抛错"到"能出报告" | 小 |
| **P1 正确** | 写前条件矩阵硬校验 + 写后完整性检查 + 缺失归因（照 `EMI-Report.py:1566-1645,1768-1854` 的语义） | 杜绝"静默写歪" | 中 |
| **P2 顺手** | 描述审计（移植 `desc_audit.py` 的规则）、图片自动匹配（移植 `picmatch.py`）、序列表格与覆盖、状态记忆 | 达到"和 EMID 差不多"的体验 | 中 |
| **P3 系统化** | `EMIReportModel` 接 `ReportGenerationService`；计划右键可一键；可选批处理/定时 | 与其它三类报告同构，可被系统复用 | 中 |

### 9.2 不建议做的事

- **不要**把 EMID 以库/嵌入方式塞进 ORTE（§5）。
- **不要**在 `EMIReportViewModel` 上继续加逻辑——它现在既是 UI 又兼 IO 又兼业务，加一层就更难拆。
- **不要**同时维护"docx 解析"和"PDF 解析"两条主链路：确定一个权威源（建议 PDF：实验室的原始产物就是 PDF），另一条只做兜底。

### 9.3 若短期只能做一件事

修 §3.1 + §3.3，并给 `DoReport` 补 `try/catch`：① 让模板选择"显式化 + 校验"（选不到就明确报错，不要回退到 Thermal Shock 模板）；② 修正 `else if` 链与 `infos` 长度检查；③ 把 `Setup` 读失败、模板缺表统一转成用户能懂的提示。这三处是"点下去就出错/出错看不出原因"的地方，改动最小、收益最直接。

### 9.4 验收基准

不管是重建（D）还是轻集成（5.4），都用**同一批真实数据**做对比验收：同一数据目录 + 同一模板，分别跑 EMID 与 ORTE，逐单元格比对 `Conducted EMI` 表（数值、合并区、行高、边框、图片锚点、备注、序列号块），差异为零或差异可解释才算通过。这个基线能同时防住"重写跑偏"与"两边行为漂移"。

---

## 10. 待确认（影响结论的 5 个问题）

1. **现场实际用的是哪套 EMI 模板？** 两条路都有问题：随包那份没有 `Setup` 表（必须先「设置→另存」才生得出来），EMID 的 `template/*.xlsx` 又是 `Setup!B1=44`（非 JSON，`ParseJson` 会抛异常）。需要确认现场到底是怎么点过「制作」的。
2. **EMI 的原始产物是 PDF 还是 docx？** ORTE 现在读 docx、打包 PDF；EMID 读 PDF。若两者都有，谁是权威？
3. **ORTE 的 EMI 需要覆盖哪些负载档数？** 1/3/4/10/AMZ/Vdc 全要，还是常用几档就够？
4. **"辅助功能"的边界**：是否需要无人值守/批量出报告？如果需要，P3（模型化）就不是可选项。
5. **EMID 与 ORTE 的长期分工**：EMID 继续做主工具、ORTE 只登记，还是 ORTE 逐步取代？这决定是"重建领域层"（D）还是"轻集成"（5.4）。
