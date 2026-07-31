# Oh My Pi (OMP) 零基础使用指南 🚀

> **欢迎使用 OMP！** 这是一个基于 AI 的命令行编程助手，它能理解你的自然语言指令，自动调用工具（如读取文件、执行代码、搜索网络等）来帮助你完成开发任务。本指南将从零开始，带你一步步掌握 OMP 的核心用法。

---

## 📋 目录

1. [快速入门](#1-快速入门)
2. [核心交互模式](#2-核心交互模式)
3. [常用快捷键与命令](#3-常用快捷键与命令)
4. [配置与管理](#4-配置与管理)
5. [高级功能：MCP 与 Extension](#5-高级功能mcp-与-extension)
6. [常见问题解答](#6-常见问题解答)

---

## 1. 快速入门

### 1.1 启动 OMP
在终端（Terminal/CMD/Git Bash）中输入以下命令即可启动交互式会话：

```bash
omp
```

启动后，你会看到一个类似聊天界面的提示符，等待你输入指令。

### 1.2 发送第一条指令
直接输入你想让 AI 做的事情，然后按 `Enter`。例如：

```text
帮我列出当前目录下的所有 Python 文件
```

AI 会自动调用 `find` 或 `bash` 工具来执行这个任务，并返回结果。

### 1.3 退出 OMP
在会话中输入 `/quit` 或按 `Ctrl+D` 即可退出。

---

## 2. 核心交互模式

### 2.1 自然语言对话
OMP 的核心是**自然语言理解**。你不需要学习复杂的命令语法，只需要像和人说话一样描述你的需求。

*   ✅ **好例子**："帮我分析一下订单.csv 文件，看看哪个产品的销售额最高。"
*   ❌ **坏例子**："执行 select max(销售额) from 订单.csv"（除非你明确想让 AI 写 SQL）

### 2.2 工具调用与审批
OMP 拥有强大的工具集（读文件、写代码、跑 bash、搜网页等）。为了安全，你可以控制 AI 使用这些工具的权限：

*   **Yolo 模式（默认）**：AI 自动执行所有操作，无弹窗干扰。适合快速探索。
*   **Write 模式**：只读操作（read/find）自动执行；危险操作（bash/browser）会弹窗询问。
*   **Always-Ask 模式**：所有操作都需确认。

**如何切换模式？**
如果你安装了 `approval-switcher` Extension（见第 5 节），只需输入：
```
/approval
```
然后选择你想要的模式即可。

### 2.3 思考级别 (Thinking Level)
OMP 有不同的"思考深度"，影响 AI 回答的详细程度和耗时：
*   `minimal` / `low`：快速响应，适合简单任务。
*   `medium` / `high`：深入分析，适合复杂问题。
*   `xhigh`：极致推理，适合架构设计或调试疑难杂症。

**如何切换？**
按 `Shift+Tab`（默认快捷键）循环切换。

---

## 3. 常用快捷键与命令

### 3.1 必知快捷键

| 快捷键 | 功能 | 说明 |
|--------|------|------|
| `Ctrl+P` | 切换模型 | 在不同 AI 模型间循环切换 |
| `Shift+Tab` | 切换思考级别 | 改变 AI 的思考深度 |
| `Alt+Shift+P` | 切换 Plan 模式 | 开启后 AI 优先输出规划而非代码 |
| `Ctrl+T` | 显示/隐藏思考块 | 控制是否展示 AI 的推理过程 |
| `Ctrl+R` | 搜索历史 | 快速查找之前的对话 |
| `Alt+H` | 语音输入 (STT) | 通过麦克风输入文字（需在 config 中启用） |

### 3.2 常用斜杠命令 (Slash Commands)

在输入框中以 `/` 开头的命令：

| 命令 | 功能 |
|------|------|
| `/quit` | 退出当前会话 |
| `/hotkeys` | 查看当前所有快捷键绑定 |
| `/model` | 选择或切换 AI 模型 |
| `/tools` | 查看当前可用的工具列表 |
| `/mcp` | 管理 MCP 服务器（见第 5 节） |
| `/approval` | 切换审批模式（需安装 Extension） |

---

## 4. 配置与管理

### 4.1 配置文件位置
OMP 的全局配置存储在：
*   **Windows**: `C:\Users\<用户名>\.omp\agent\config.yml`
*   **macOS/Linux**: `~/.omp/agent/config.yml`

### 4.2 常用配置项
你可以直接用文本编辑器打开 `config.yml` 进行修改，或使用 `omp config` 命令。

**示例：启用语音输入 (STT)**
```bash
omp config set stt.enabled true
omp config set stt.modelName base.en
omp config set stt.language zh
```

**示例：更改主题**
```bash
omp config set theme.dark monokai
```

### 4.3 查看与修改配置
*   **列出所有配置**：`omp config list`
*   **查看单项配置**：`omp config get <key>`
*   **设置配置**：`omp config set <key> <value>`

---

## 5. 高级功能：MCP 与 Extension

### 5.1 MCP (Model Context Protocol)
MCP 允许 OMP 连接外部数据源或工具（如 Excel、数据库、GitHub 等）。

**如何配置 MCP 服务器？**
1.  创建或编辑用户级 MCP 配置文件：`~/.omp/agent/mcp.json`
2.  添加服务器定义，例如：
    ```json
    {
      "mcpServers": {
        "excel-pivot": {
          "command": "C:\\Users\\shifei2\\.excelpivot\\ExcelPowerPivotMcp.exe",
          "args": []
        }
      }
    }
    ```
3.  重启 OMP。
4.  在会话中输入 `/mcp list` 确认连接成功。
5.  直接对 AI 说："请用 excel-pivot 帮我分析订单数据"，AI 会自动调用该工具。

### 5.2 Extension (扩展)
Extension 是自定义的 TypeScript/JavaScript 模块，可以增强 OMP 的功能（如自定义快捷键、拦截工具调用、添加新命令等）。

**如何安装 Extension？**
1.  将 Extension 文件夹（包含 `index.ts`）复制到：`~/.omp/agent/extensions/`
2.  重启 OMP。
3.  Extension 会自动加载。

**示例：Approval Switcher**
本目录下已包含一个 `approval-switcher` Extension，它提供了 `/approval` 命令，让你能在会话内实时切换审批模式，无需重启。

---

## 6. 常见问题解答

### Q1: OMP 不识别我的中文指令怎么办？
A: OMP 原生支持多语言。如果效果不佳，可以尝试在 `config.yml` 中设置系统提示词，或确保你的模型支持中文（如 Qwen 系列）。

### Q2: 如何重置 OMP 到初始状态？
A: 删除 `~/.omp/agent/` 目录下的所有文件（注意备份重要配置），然后重新启动 OMP。

### Q3: MCP 连接失败怎么办？
A: 
1. 检查 `mcp.json` 中的路径是否正确。
2. 确保 MCP Server 的可执行文件有运行权限。
3. 在 OMP 中输入 `/mcp test <server-name>` 进行诊断。

### Q4: 如何分享我的 OMP 配置给别人？
A: 打包 `~/.omp/agent/` 目录（排除 `sessions/` 和 `blobs/` 等大文件夹），发送给对方解压即可。

---

## 🎉 结语

OMP 是一个强大且灵活的 AI 编程伙伴。随着你使用的深入，你会发现更多有趣的功能和技巧。祝你使用愉快！

如有任何问题，欢迎查阅 OMP 官方文档或社区论坛。
