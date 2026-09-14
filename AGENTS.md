# ORT一键报告 — 仓库协作约定

## 每轮对话收尾：提交并同步（强制）

- 一轮对话的所有文件变更完成后、**给出最终回复之前**，必须加载 skill `git-commit-sync`（`.dsh/skills/git-commit-sync/SKILL.md`）并按其步骤执行：
  暂存全部变更 → 按仓库既有中文 Conventional Commits 风格编写 commit → `git commit` → `git push origin <当前分支>`。
- 只有本轮**没有任何文件改动**时才跳过，并在回复中写一行「本轮无文件变更，未提交」。
- 用户明确要求「不要提交 / 不要推送」时不执行。

## 构建与验证

- 构建：`MSBuild .\ORT一键报告.csproj /p:Configuration=Debug /p:SignManifests=false /p:GenerateManifests=false /v:minimal /nologo /t:Rebuild`（清单签名在受限环境会失败，与代码无关）。
- 重建 `bin\Debug\*` 前先确认没有 `ORT一键报告.exe` 在运行，否则 `SQLite.Interop.dll` 被占用。
- 数据库为 SQLite + FreeSql（`UseAutoSyncStructure(true)`），新增实体列会自动迁移。
