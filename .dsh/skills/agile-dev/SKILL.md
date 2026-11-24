---
name: agile-dev
description: "敏捷开发模式：以最快速度完成一轮对话——只要编译通过，不跑单元测试、不写探针/冒烟程序；不提交、不推送、不写更新日志、不升版本号，也不做额外收尾流程；直到用户明确说『提交 / 同步 / 发布 / 更新日志』才做那些。用户说『敏捷开发』『快速改』『别测试别提交』『尽快完成这轮』时启用。"
whenToUse: "用户要求敏捷开发、尽快完成一轮对话、只要求编译通过、不要测试/不要提交/不要更新日志时；或用户说『先这样，别提交』。"
---

# 敏捷开发模式（ORT一键报告）

## 启用条件

用户说「敏捷开发」「快速改一下」「别测试别提交」「尽快完成这轮」等 → 本轮起按本技能执行。与 `AGENTS.md` 的强制收尾流程（`git-commit-sync`、`update-log`）冲突时，**用户指令与本技能优先**。

## 硬规则

1. **只要求编译通过**：改完代码跑一次构建即可。
   - `MSBuild .\ORT一键报告.csproj /p:Configuration=Debug /p:SignManifests=false /p:GenerateManifests=false /v:minimal /nologo /t:Build`（全量用 `/t:Rebuild`；VS 的 MSBuild 路径可用 `vswhere` 查）。
   - **不跑** `Tests\run-tests.ps1`、不写临时验证/探针/冒烟程序、不做界面点击验证。
   - 编译报错 → 修到能过再收尾；警告不处理。exe 在运行导致复制失败（`SQLite.Interop.dll` / `ORT实验室管理系统.exe` 被占用）时，说明一句即可，不必反复重试。
2. **不提交、不推送**：本轮不 `git add` / `commit` / `push`，改动留在工作区。用户明确说「提交」「存一下」「同步」时才按 `git-commit-sync` 走。
3. **不写更新日志、不升版本号**：不碰 `更新日志.md`，不动 `Properties/AssemblyInfo.cs` 与 csproj 的版本号。用户明确说「更新日志」「发布 X.Y」时才按 `update-log` 走。
4. **不做额外流程**：不建 todo 清单、不写方案文档、不铺陈推理与备选方案、不改与需求无关的文件、不做顺手重构。需求有歧义就按最合理理解直接做，做完点明假设，不停下来追问。
5. **开局自检从简**：`git-status-check` 只确认分支与工作区（可省 `fetch`），一行带过。
6. **收尾一句话**：例如「已改完，编译通过，未提交（敏捷模式）」，等用户下一步。

## 边界（仍然要守）

- 改动尽量局部，不破坏既有功能。
- 不删用户未提交的改动；绝不 `git reset --hard` / `checkout -- .` / `clean` / `push --force`。
- 用户要求测试、提交或发布时，立刻回到常规流程（`git-commit-sync`、`update-log`、`Tools\make-release.ps1`）。
