---
name: git-status-check
description: "本仓库每轮对话开局的自检流程：动手之前先查 Git 状态——当前分支、工作区是否干净、本地最新提交、以及 origin 上有没有我们还没看到的新提交。任何一轮对话在读取代码或修改文件之前都应执行，避免在过期的代码基线上工作、避免覆盖上一轮的遗留改动；用户问『现在改到哪了』『有没有最新更改』『同步了吗』时同样走本流程。"
whenToUse: "每当一轮对话开始、准备读代码或改文件之前；或用户询问当前仓库/分支状态、最新改动、是否与远端同步时。"
---

# 每轮开局：检查 Git 状态（ORT一键报告）

## 何时执行

- **默认执行**：每轮对话收到第一条实质指令后、**动任何文件之前**先走一遍下面的步骤。AGENTS.md 里已登记这条强制约定。
- **必须执行**：用户问「现在改到哪了 / 有没有最新更改 / 和远端同步了吗 / 上一轮改了什么」。
- **跳过**：与仓库无关的纯问答、只是想聊两句 —— 可不查，但不要假装查过。

## 步骤

全部在仓库根目录执行（用 `workdir` 参数，**不要 `cd`**）。

### 1. 认清仓库与分支

- `git rev-parse --show-toplevel` —— 确认确实在 ORT一键报告 仓库内。
- `git rev-parse --abbrev-ref HEAD` —— 当前分支；输出 `HEAD` 表示 detached，需在汇报里点明。

### 2. 看工作区

- `git status --short --branch` —— 一次拿到分支 + 文件清单 + 与上游的 ahead/behind。
- 有未提交改动时先判断来源：上一轮遗留、还是用户自己手改的。**只报告，不清场** —— 后续工作要在这批改动之上继续，不能当它不存在。

### 3. 看本地最新提交

- `git log --oneline -5` —— 掌握最近做了什么（也是写「更新内容」时的上下文）。

### 4. 看远端有没有新东西（需要网络）

- `$env:GIT_TERMINAL_PROMPT='0'; git fetch origin <当前分支> --prune`
  - 受限沙箱下凭据助手可能起不来，报错形如 `sh.exe: *** fatal error - couldn't create signal pipe` 或 `fatal: could not read Username for 'https://github.com'`。这不是仓库坏了，用 `sandbox_permissions: danger-full-access` 重跑**同一条** `git fetch`，justification 写「Git 凭据助手在受限沙箱下无法创建命名管道，需要它完成远端状态查询」。
  - 网络偶发失败（本机走 127.0.0.1:7897 代理，常见 `SSL_ERROR_SYSCALL`）：重试一次；仍失败就跳过这步，在汇报里写明「未取到远端状态」。
- `git status -sb` —— 看 `[ahead N]` / `[behind N]`。
- `git log --oneline HEAD..origin/<当前分支>` —— 列出**远端有、本地没有**的新提交（真正意义上的「最新更改」）。

### 5. 落后远端时

- **只报告，不自动同步**：不擅自 `git pull` / `git merge` / `git rebase` / `git stash`，这些都可能在用户有未提交改动时造成惊喜。
- 用户明确要求同步时才执行，并按 skill `git-commit-sync` 里的冲突与推送处理规则办。

### 6. 汇报（一到三行，不刷屏）

固定给出这几项：分支｜工作区｜相对远端｜最近提交 / 需要注意的遗留改动。示例：

- `分支 master｜工作区干净｜与 origin/master 一致｜最近提交 92bf456 fix(plans): 修复筛选合并项点不开…`
- `分支 master｜3 个未提交改动（计划表筛选相关）｜落后 origin/master 2 个提交：a1b2c3d …、d4e5f6a …`
- `分支 master｜落后远端但 fetch 失败（沙箱凭据）｜可手动执行 git fetch origin master`

同一轮内只在开局报一次；之后若发生实质变化（自己提交过、发现远端新提交）再补报。

## 硬性约束

- 本流程**只读**：检查阶段绝不 `add` / `commit` / `pull` / `merge` / `rebase` / `reset` / `checkout` / `clean` / `stash pop`，不改动工作区。
- 绝不 `git reset --hard`、`git checkout -- .`、`git clean` —— 会毁掉用户未提交的工作。
- fetch/网络失败不阻塞本轮工作：如实说明 + 给出手动命令，然后继续。
- 发现分支为 detached、或工作区有大量来源不明的改动时，先向用户确认再动手。
