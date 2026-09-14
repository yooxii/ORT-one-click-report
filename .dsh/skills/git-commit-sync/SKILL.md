---
name: git-commit-sync
description: "本仓库每轮对话收尾的强制流程：一轮对话的所有文件变更完成后，暂存全部变更、按仓库既有风格编写中文 commit、提交并推送到 origin（同步）。任何一轮只要产生过文件改动（新增/修改/删除/重命名，含由脚本或构建产生的受版本控制文件），在给出最终回复之前都必须执行；只有本轮完全无改动才可跳过。用户直接说『提交』『存一下』『同步到远端』时同样走本流程。"
whenToUse: "每当一轮对话完成全部文件变更、准备给出最终回复时；或用户要求把当前改动提交/同步时。"
---

# 每轮收尾：提交并同步（ORT一键报告）

## 何时执行

- **默认执行**：本轮只要改过任何文件，就在写最终回复**之前**执行下面的流程。
- **跳过**：本轮完全无文件改动 —— 在回复中写一行「本轮无文件变更，未提交」。
- **不执行**：用户明确说「先别提交 / 不要推送 / 我自己来」时，用户指令优先。

## 步骤

### 1. 确认仓库与状态

在仓库根目录执行（用 `workdir` 参数，不要 `cd`）：

- `git rev-parse --abbrev-ref HEAD` —— 记下当前分支名，推送时要用。
- `git status --porcelain` —— 输出为空即无变更，流程结束。
- 分支名为 `HEAD`（detached）时不要提交，直接报告并停止。

### 2. 暂存全部变更

- `git add -A`
- 禁止 `git add -f`。被 `.gitignore` 忽略的构建产物（`bin/`、`obj/`、`*.log`、`*.pfx`、`*.user`…）留在工作区即可。
- 用 `git diff --cached --stat` 复核清单。若出现不该入库的内容（密钥/令牌/口令/证书/本地数据库/临时验证脚本/截图），先取消暂存并向用户说明，不要硬提交。

### 3. 编写 commit message

遵循仓库既有约定（可用 `git log -3 --format='%B'` 复核）：**中文**、Conventional Commits 前缀 + 可选 scope。

- 标题：`<type>(<scope>): 一句话说明本轮做了什么`，≤ 50 字，结尾不加句号。
  - type：`feat` 新功能 / `fix` 修缺陷 / `refactor` 重构 / `perf` 性能 / `style` 界面样式 / `docs` 文档 / `chore` 杂务 / `test` 测试 / `build` 构建
  - scope 用模块名，例如 `mail`、`admin`、`review`、`theme`、`i18n`、`plans`、`ui`
- 正文：空一行后，每条 `- ` 开头一行，逐条说明**实际改了什么**（功能点、关键文件、行为变化、修复的问题），不写空话。
- 不添加署名或工具尾注（仓库历史中不存在）。

默认一轮一次提交；当变更明显属于多个互不相关的逻辑单元时，可拆成多个提交（按路径逐个 `git add -- <paths>` 后分别提交）。

### 4. 提交

中文正文与多行内容在 PowerShell 里极易被引号与编码破坏，**必须走临时文件**：

1. 用 write 工具把完整 message 写成临时文件（UTF-8 **无 BOM**），放在 `.git/COMMIT_MSG_TMP`、`$env:TEMP` 或仓库内被忽略的 `obj/` 下。
2. `git commit -F <临时文件>`
3. 在同一个 `git add -A` 之后删除该临时文件，确保它不会进入工作区或索引。

提交因钩子/校验失败时，读错误、修正后重新提交；**不得**用 `--no-verify` 绕过，也不得改写已有历史。

### 5. 同步（推送）

- `git push origin <分支名>`，分支名取自步骤 1。
- 必须非交互：设 `$env:GIT_TERMINAL_PROMPT='0'`，避免凭据弹窗把命令挂死。
- 被拒（non-fast-forward）：`git fetch origin` → `git rebase origin/<分支名>` → 解决后重试推送；冲突无法安全解决就 `git rebase --abort`，保留本地提交并如实报告。
- 网络失败（本机走 127.0.0.1:7897 代理，偶发 `SSL_ERROR_SYSCALL`）：**重试一次**；仍失败则保留本地提交，并在回复中给出用户可手动执行的命令。
- **受限沙箱下凭据助手无法启动**（本机 Git 凭据由 Git Credential Manager 保管）：报错形如 `sh.exe: *** fatal error - couldn't create signal pipe, Win32 error 5` 加 `fatal: could not read Username for 'https://github.com'`。这**不是**凭据缺失，普通重试无用 —— 直接以 `sandbox_permissions: danger-full-access` 重跑**同一条** push 命令，justification 写「Git 凭据助手在受限沙箱下无法创建命名管道，需要其完成推送」，可一次成功。
- 推送完成后用 `git status --porcelain` + `git log --oneline -1` 复核工作区干净、本地与远端一致。

### 6. 汇报

最终回复里固定给出一到两行结果，例如：

- `已提交：feat(mail): 新增邮件通知服务（a1b2c3d），已推送 origin/master`
- `已提交（本地）：… ；推送失败：<原因>，可手动执行 git push origin master`
- `本轮无文件变更，未提交`

## 硬性约束

- 绝不 `git push --force` / `-f`；绝不 `git reset --hard`、`git checkout -- .`、`git clean`（会毁掉用户未提交的工作）。
- 绝不 `--amend` 已推送的提交，不 rebase 已推送的历史。
- 不提交密钥、令牌、口令、证书（`*.pfx`）、数据库、`bin`/`obj` 产物、临时验证脚本与截图。
- 一轮之内若在提交后又产生了新改动，收尾时再补一次「add → commit → push」。
