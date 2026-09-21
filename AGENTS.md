# ORT一键报告 — 仓库协作约定

## 每轮对话收尾：提交并同步（强制）

- 一轮对话的所有文件变更完成后、**给出最终回复之前**，必须加载 skill `git-commit-sync`（`.dsh/skills/git-commit-sync/SKILL.md`）并按其步骤执行：
  暂存全部变更 → 按仓库既有中文 Conventional Commits 风格编写 commit → `git commit` → `git push origin <当前分支>`。
- 只有本轮**没有任何文件改动**时才跳过，并在回复中写一行「本轮无文件变更，未提交」。
- 用户明确要求「不要提交 / 不要推送」时不执行。

## 每轮对话收尾：更新内容与更新日志（按需）

- 每轮有文件改动，先加载 skill `update-log`（`.dsh/skills/update-log/SKILL.md`）把本轮改动总结成「更新内容」。
- 当用户声明「这是一次版本更新」，或该轮**程序改动**大到足以当作一次更新（跨多模块、有面向用户的新功能/行为变化等，判定标准见该 skill）时，在**根目录 `更新日志.md`** 追加一节：正序追加，最新小节在文件末尾，同日并入末尾小节，不改写历史条目。
- `更新日志.md` **只写程序本身的更新内容**（功能/行为/界面变化、缺陷修复、数据与设置格式变化、破坏性变化）；技能、`AGENTS.md` 约定、构建与验证脚本、纯文档等**仓库协作类改动不写入**。
- `修改记录.md` 的用法保持原样：本约定不为它新增规则，也不改写、删除其既有内容。
- 顺序固定：先写更新日志 → 再走上面的 `git-commit-sync`，让日志进入同一个 commit。

## 构建与验证

- 构建：`MSBuild .\ORT一键报告.csproj /p:Configuration=Debug /p:SignManifests=false /p:GenerateManifests=false /v:minimal /nologo /t:Rebuild`（清单签名在受限环境会失败，与代码无关）。
- 重建 `bin\Debug\*` 前先确认没有 `ORT一键报告.exe` 在运行，否则 `SQLite.Interop.dll` 被占用。
- 数据库为 SQLite + FreeSql（`UseAutoSyncStructure(true)`），新增实体列会自动迁移。
- 发布包自带更新日志：csproj 里已把根目录 `更新日志.md` 登记为 `CopyToOutputDirectory=PreserveNewest`，生成时自动复制到 `bin\Debug\`（Release 同理）。发布前先写完 `更新日志.md` 再生成，并用文件哈希核对输出目录里的日志与根目录一致；不要移除这条 csproj 登记。
