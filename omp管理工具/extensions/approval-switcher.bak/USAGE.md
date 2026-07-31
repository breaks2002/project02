# OMP 审批模式切换扩展 (Approval Switcher) — 升级版

## 📦 功能简介
允许你在**不重启会话**的情况下，通过 `/approval` 命令或 **F2 快捷键**实时切换工具的审批模式。

### 支持的四种模式：

| 模式 | 行为 | 状态标识 |
|---|---|---|
| **default** | 每个 write/edit/bash 都弹窗确认 | ⏵ |
| **plan** | 只读；仅允许 read/search + 安全 bash 白名单 | ⏸ |
| **acceptEdits** | write/edit 静默放行，bash 弹窗确认 | ⏵⏵ |
| **bypassPermissions** | 全部放行（仅拦截灾难性命令和受保护路径） | ⏵⏵⏵⏵ |

---

## 🛡️ 始终生效的安全底线（与模式无关）

以下检查**在任何模式下都会执行**，包括 bypassPermissions：

1. **灾难性命令拦截**：`mkfs`、`dd`、fork bomb、写 `/dev/sda` 等
2. **rm -rf 关键目录检测**：阻止删除 `/`、`/etc`、`/bin`、`/usr` 等系统目录
3. **受保护路径**：`~/.ssh`、`~/.aws`、`~/.kube/config`、`~/.npmrc` 等敏感配置禁止修改
4. **Session 级审批缓存**：在 default/acceptEdits 模式下，一次 approve 后同工具/命令不再重复弹框

---

## 🚀 使用方法

- **按 F2** → 循环切换模式：default → plan → acceptEdits → bypassPermissions → default
- **`/approval`** → 弹出交互式选择菜单
- **`/approval default`** → 直接切换到每次确认
- **`/approval plan`** → 直接切换到只读模式
- **`/approval acceptEdits`** → 直接切换到半自动
- **`/approval bypassPermissions`** → 直接切换到全自动

---

## 📋 Plan 模式详解

Plan 模式适合**先探索代码再动手**的场景：

**放行的工具**：read、search、find、lsp、ast_grep

**放行的 bash 命令**（约 60+ 个只读前缀）：
- 文件查看：`cat`、`head`、`tail`、`less`、`wc`、`stat`
- 搜索：`grep`、`find`、`rg`、`fd`
- Git 只读：`git status`、`git log`、`git diff`、`git branch`
- 系统信息：`ps`、`uname`、`env`、`date`
- 包管理：`npm list`、`npm view`、`npm audit`

**阻止的操作**：write、edit、非安全 bash 命令、browser、task、debug

---

## 🔧 工作原理

- **零依赖**：纯 TypeScript，仅依赖 OMP 内置 SDK
- **实时生效**：通过 `tool_call` 事件动态决策
- **三层安全检查**：
  1. 始终生效的安全层（灾难性命令 + 受保护路径 + rm-rf 检测）
  2. Plan 模式的工具级只读约束
  3. 模式级的审批逻辑（default / acceptEdits / bypassPermissions）
- **Session 审批缓存**：default/acceptEdits 模式下，首次确认后会话内同操作不再弹窗

---

## ❓ 常见问题

**Q: 为什么我输入 `/approval` 没反应？**
A: 请确认已重启 OMP，且文件夹路径正确。检查 `~/.omp/agent/extensions/` 下是否有该文件夹。

**Q: 弹窗点"拒绝"后会发生什么？**
A: 工具调用会被阻断，AI 会收到"用户拒绝"的反馈。

**Q: bypassPermissions 模式下，AI 能删除系统目录吗？**
A: 不能。`rm -rf /`、`rm -rf /etc` 等关键目录删除操作在任何模式下都会被拦截。

**Q: plan 模式和 acceptEdits 模式的区别？**
A: plan 模式是**真正的只读**——write/edit/bash 等写工具被阻止，AI 只能 read/search。acceptEdits 模式是半自动——所有工具可用，但执行类工具需要确认。
