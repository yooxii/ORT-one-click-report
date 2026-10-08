# ORT一键报告 — 仓库协作约定

## 每轮对话开局：检查 Git 状态（强制）

- 每轮对话收到第一条实质指令后、**读取或修改任何文件之前**，必须加载 skill `git-status-check`（`.dsh/skills/git-status-check/SKILL.md`）并按其步骤执行：
  确认分支 → 看工作区 → 看本地最新提交 → `git fetch origin <当前分支>` 后比对远端新提交 → 一到三行汇报（分支｜工作区｜与远端关系｜最近提交）。
- 本步骤**只读**：不自动 `pull` / `merge` / `rebase`，不清理用户未提交的改动；落后远端只报告，用户明确要求同步时才执行。
- `git fetch` 因沙箱凭据或网络失败时不阻塞本轮工作：如实说明并给出手动命令，然后继续。
- 与仓库无关的纯问答可跳过，但不得假装查过。

## 每轮对话收尾：提交并同步（强制）

- 一轮对话的所有文件变更完成后、**给出最终回复之前**，必须加载 skill `git-commit-sync`（`.dsh/skills/git-commit-sync/SKILL.md`）并按其步骤执行：
  暂存全部变更 → 按仓库既有中文 Conventional Commits 风格编写 commit → `git commit` → `git push origin <当前分支>`。
- 只有本轮**没有任何文件改动**时才跳过，并在回复中写一行「本轮无文件变更，未提交」。
- 用户明确要求「不要提交 / 不要推送」时不执行。

## 每轮对话收尾：更新内容与更新日志（按需）

- 每轮有文件改动，先加载 skill `update-log`（`.dsh/skills/update-log/SKILL.md`）把本轮改动总结成「更新内容」。
- **版本发布由用户主动提出**（「这是一次版本更新 / 发布 0.6 / 记一下这一版」）：**不发布**任何单纯按改动量推断的版本，也不按「条目数达到某个阈值」自动新开版本小节。用户没提，就只做「更新内容」摘要、不写 `更新日志.md`。
- 用户提出发布时，在**根目录 `更新日志.md`** 追加一节**版本**记录：标题为 `## <版本号>`（如 `## 0.5`，不用日期），新版追加在文件末尾；**版本号由用户指定**（用户没给就按用户措辞里的号，仍不确定才问一句，不要自行 +0.1）。
- **版本号必须三处一致**：`更新日志.md` 的版本小节、`Properties/AssemblyInfo.cs` 的 `AssemblyVersion`/`AssemblyFileVersion`、csproj 的 `ApplicationVersion`（主界面显示的版本号取自程序集版本，不一致就会被用户看到）。用户要求改版本号时三处一起改。
- `更新日志.md` **只写程序功能与行为变化**（新功能、行为/界面变化、缺陷修复、数据与设置格式变化、破坏性变化）；**不写实现细节**（文件/类/方法名、代码清理与重构、单元测试、依赖与构建、提交与验证过程），技能、`AGENTS.md` 约定、脚本、纯文档等**仓库协作类改动也不写入**。
- 写入前先查重压缩：同一功能只保留一条、多次演进只保留最终形态，能合并就合并；用户明确要求「统一到某一版」时，把之后的版本小节内容并入该版本并删掉多余标题（保留全部功能要点）。
- `修改记录.md` 的用法保持原样：本约定不为它新增规则，也不改写、删除其既有内容。
- 顺序固定：先写更新日志 → 再走上面的 `git-commit-sync`，让日志进入同一个 commit。

## 发布打包（用户提出发布时执行）

- 打包脚本：`tools/make-release.ps1`（PowerShell，含中文，**必须保留 UTF-8 BOM**，否则 PowerShell 5.1 按 ANSI 读会语法报错）。
- 脚本行为：重新生成 Release（`-SkipBuild` 可跳过）→ 把 Release 里运行必需的文件按原目录结构暂存 → 压缩为 `bin\ORT实验室管理系统_<主.次>.zip`，包内顶层目录同名 → 排除 `logs/`、`*.pdb`、`*.xml` → 自检 `Data\local_settings.json` 与 `更新日志.md` 是否在包内、日志哈希是否与根目录一致。
- 版本号从 `Properties\AssemblyInfo.cs` 读取（匹配行首的 `[assembly: AssemblyVersion(...)]`，注意文件里还有一行被注释的示例 `1.0.*`）。
- 打包前先确认没有 `ORT实验室管理系统.exe` 在运行，否则 Release 生成会在覆盖 exe 时失败。
- zip 落在 `bin\`（该目录已被 `.gitignore` 忽略，不进版本库）。

## 构建与验证

- 构建：`MSBuild .\ORT一键报告.csproj /p:Configuration=Debug /p:SignManifests=false /p:GenerateManifests=false /v:minimal /nologo /t:Rebuild`（清单签名在受限环境会失败，与代码无关）。
- 重建 `bin\Debug\*` 前先确认没有 `ORT实验室管理系统.exe` 在运行，否则 `SQLite.Interop.dll` 被占用。（输出 exe 名 = 程序名 `ORT实验室管理系统`；csproj 文件名与根命名空间仍为 `ORT一键报告`。）
- 数据库为 SQLite + FreeSql（`UseAutoSyncStructure(true)`），新增实体列会自动迁移。
- 单元测试：`Tests\run-tests.ps1`（按字节加载程序集，受限环境也能跑）。
- 发布包自带更新日志：csproj 里已把根目录 `更新日志.md` 登记为 `CopyToOutputDirectory=PreserveNewest`，生成时自动复制到 `bin\Debug\`（Release 同理）。发布前先写完 `更新日志.md` 再生成，并用文件哈希核对输出目录里的日志与根目录一致；不要移除这条 csproj 登记。
