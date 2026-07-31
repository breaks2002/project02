import type { ExtensionAPI } from "@oh-my-pi/pi-coding-agent";
import { homedir } from "node:os";
import { resolve } from "node:path";

// ─── 模式定义 ───────────────────────────────────────────────────────────

const MODES = ["default", "plan", "acceptEdits", "bypassPermissions"] as const;
type Mode = (typeof MODES)[number];

let currentMode: Mode = "bypassPermissions";

const MODE_LABELS: Record<Mode, string> = {
  default: "default（每个 write/edit/bash 都弹窗确认）",
  plan: "plan（只读 — 仅允许 read/search + 安全 bash 白名单）",
  acceptEdits: "acceptEdits（write/edit 静默放行，bash 弹窗确认）",
  bypassPermissions: "bypassPermissions（全部放行，仅拦截灾难性命令和受保护路径）",
};


// write 模式下需要确认的执行类工具
const EXEC_TOOLS: Record<string, true> = {
  bash: true,
  browser: true,
  task: true,
  debug: true,
};

// plan 模式允许的只读工具 + 安全 bash 白名单
const PLAN_TOOLS: Record<string, true> = {
  read: true,
  search: true,
  find: true,
  lsp: true,
  ast_grep: true,
};

const SAFE_PLAN_BASH_PREFIXES = [
  "cat", "head", "tail", "less", "more", "grep", "find", "ls",
  "pwd", "echo", "printf", "wc", "sort", "uniq", "diff", "file",
  "stat", "du", "df", "tree", "which", "whereis", "type", "env",
  "printenv", "uname", "whoami", "id", "date", "cal", "uptime",
  "ps", "top", "htop", "free", "curl", "jq", "sed", "awk",
  "rg", "fd", "bat", "eza", "git status", "git log", "git diff",
  "git show", "git branch", "git remote", "git ls-", "git config --get",
  "npm list", "npm ls", "npm view", "npm info", "npm search",
  "npm outdated", "npm audit",
];

// ─── 始终生效的安全层 ───────────────────────────────────────────────────

interface Pattern {
  pattern: string;
  description: string;
}

const CATASTROPHIC_PATTERNS: Pattern[] = [
  { pattern: "sudo mkfs", description: "sudo 格式化文件系统" },
  { pattern: "mkfs.", description: "格式化文件系统" },
  { pattern: "dd if=", description: "原始磁盘写入" },
  { pattern: ":(){ :|:& };:", description: "fork bomb" },
  { pattern: "> /dev/sda", description: "覆盖磁盘" },
  { pattern: "> /dev/nvme", description: "覆盖磁盘" },
  { pattern: "sudo dd", description: "sudo 原始磁盘操作" },
];

const DANGEROUS_PATTERNS: Pattern[] = [
  { pattern: "chmod -R 777", description: "不安全的递归权限" },
  { pattern: "chown -R", description: "递归所有权变更" },
  { pattern: "> /dev/", description: "直接设备写入" },
];

const CRITICAL_DIRS = [
  "/", "/bin", "/boot", "/dev", "/etc", "/home", "/lib", "/lib64",
  "/opt", "/proc", "/root", "/run", "/sbin", "/srv", "/sys",
  "/tmp", "/usr", "/var",
];

const DEFAULT_PROTECTED_PATHS = [
  "~/.ssh", "~/.aws", "~/.gnupg", "~/.gpg",
  "~/.bashrc", "~/.bash_profile", "~/.profile",
  "~/.zshrc", "~/.zprofile",
  "~/.config/git/credentials",
  "~/.netrc", "~/.npmrc",
  "~/.docker/config.json", "~/.kube/config",
];

// Session 级审批缓存
const sessionAllow: { tools: Set<string>; commands: Set<string> } = {
  tools: new Set(),
  commands: new Set(),
};

function isSessionAllowed(toolName: string, input: Record<string, unknown>): boolean {
  if (sessionAllow.tools.has(toolName)) return true;
  if (toolName === "bash") {
    const cmd = typeof input.command === "string" ? input.command : "";
    if (cmd && sessionAllow.commands.has(cmd)) return true;
  }
  return false;
}

function recordSessionApproval(toolName: string, input: Record<string, unknown>): void {
  sessionAllow.tools.add(toolName);
  if (toolName === "bash" && typeof input.command === "string" && input.command) {
    sessionAllow.commands.add(input.command);
  }
}

// ─── 安全检查函数 ───────────────────────────────────────────────────────

function resolveProtectedPath(path: string): string {
  const home = homedir();
  if (path.startsWith("~/")) {
    return resolve(home, path.slice(2));
  }
  return resolve(path);
}

function isPathProtected(targetPath: string, protectedPaths: string[]): boolean {
  return protectedPaths.some((protectedPath) => {
    if (targetPath === protectedPath || targetPath.startsWith(protectedPath + "/")) {
      return true;
    }
    // 检查精确文件匹配（如 ~/.bashrc）
    if (protectedPath.includes("~")) {
      const resolved = resolveProtectedPath(protectedPath);
      return targetPath === resolved;
    }
    return false;
  });
}
function isCriticalRmRf(command: string): string | null {
  const rmRfMatch = command.match(/rm\s+(-[a-zA-Z]*r[a-zA-Z]*f[a-zA-Z]*|-[a-zA-Z]*f[a-zA-Z]*r[a-zA-Z]*)\s+(.+)/);
  if (!rmRfMatch) return null;
  const target = rmRfMatch[3];
  if (!target) return null;
  for (const dir of CRITICAL_DIRS) {
    if (target === dir || target.startsWith(dir + "/") || target.startsWith(dir + " ")) {
      return `⚠️ 检测到 rm -rf 针对关键目录: ${dir}`;
    }
  }
  return null;
}
// ─── 主扩展 ─────────────────────────────────────────────────────────────

export default function approvalSwitcher(pi: ExtensionAPI) {
  pi.setLabel("Approval Switcher");

  const protectedPaths = DEFAULT_PROTECTED_PATHS.map((p) => resolveProtectedPath(p));
  const cwd = process.cwd();

  function updateStatus(ctx: { ui: { setStatus: (id: string, label: string, color?: string) => void } }): void {
    const label = MODE_LABELS[currentMode];
    let color = "green";
    if (currentMode === "acceptEdits") color = "yellow";
    if (currentMode === "plan") color = "blue";
    if (currentMode === "default") color = "red";

    ctx.ui.setStatus("approval-mode", label, color);
  }

  pi.on("session_start", async (_event: unknown, ctx: { ui: { setStatus: (id: string, label: string, color?: string) => void } }): Promise<void> => {
    sessionAllow.tools.clear();
    sessionAllow.commands.clear();
    currentMode = "bypassPermissions";
    updateStatus(ctx);
  });

  // /approval 命令
  pi.registerCommand("approval", {
    description: "切换审批模式（default / plan / acceptEdits / bypassPermissions）",
    handler: async (args: string, ctx: {
      ui: {
        select: (title: string, choices: Array<{ value: string; label: string; detail: string }>, type?: string) => Promise<string | undefined>;
        notify: (message: string, type?: string) => void;
        setStatus: (id: string, label: string, color?: string) => void;
      };
    }): Promise<void> => {
      const requested = args.trim();
      if (MODES.includes(requested as Mode)) {
        currentMode = requested as Mode;
        updateStatus(ctx);
        ctx.ui.notify(`已切换为: ${MODE_LABELS[currentMode]}`, "info");
        return;
      }

      const choices = MODES.map((mode) => ({
        value: mode,
        label: MODE_LABELS[mode],
      }));

      const selected = await ctx.ui.select("请选择审批模式:", choices, "info");

      if (typeof selected === "string") {
        const modeMatch = selected.match(/(default|plan|acceptEdits|bypassPermissions)/);
        if (modeMatch) {
          currentMode = modeMatch[1] as Mode;
          updateStatus(ctx);
          ctx.ui.notify(`已切换为: ${MODE_LABELS[currentMode]}`, "info");
        } else {
          ctx.ui.notify("[错误] 无法解析选中的模式", "error");
        }
      }
    },
  });

  // F2 快捷键：循环切换模式（default → plan → acceptEdits → bypassPermissions → default）
  pi.registerShortcut("f2", {
    description: "切换审批模式",
    handler: async (ctx) => {
      const idx = MODES.findIndex((m) => m === currentMode);
      currentMode = MODES[(idx + 1) % MODES.length];
      updateStatus(ctx);
      ctx.ui.notify(`已切换为: ${MODE_LABELS[currentMode]}`, "info");
    },
  });

  // 工具调用拦截
  pi.on("tool_call", async (event: { toolName: string; input: Record<string, unknown> }, ctx: {
    ui: {
      confirm: (message: string, choices: string[], type?: string) => Promise<boolean>;
      notify: (message: string, type?: string) => void;
      setStatus: (id: string, label: string, color?: string) => void;
    };
    cwd?: string;
  }): Promise<{ block: true; reason: string } | undefined> => {
    const tool = event.toolName;
    const input = event.input;

    // ═══════════════════════════════════════════════════
    // 第一层：始终生效的安全检查（与模式无关）
    // ═══════════════════════════════════════════════════

    // 1. 写/编辑工具的受保护路径检查
    if (tool === "write" || tool === "edit") {
      const targetPath = typeof input.path === "string" ? resolve(input.path) : "";
      if (targetPath && isPathProtected(targetPath, protectedPaths)) {
        return { block: true, reason: `受保护路径，禁止修改: ${targetPath}` };
      }
    }

    // 2. Bash 命令安全检查
    if (tool === "bash") {
      const command = typeof input.command === "string" ? input.command : "";

      // 灾难性命令拦截（始终阻止）
      const catastrophic = CATASTROPHIC_PATTERNS.find((p) => command.includes(p.pattern));
      if (catastrophic) {
        return { block: true, reason: `🚫 灾难性命令已阻止: ${catastrophic.description}` };
      }

      // rm -rf 关键目录检测
      const criticalRm = isCriticalRmRf(command);
      if (criticalRm) {
        return { block: true, reason: criticalRm };
      }

      // 危险命令拦截（bypassPermissions 模式放行，其他模式阻止）
      if (currentMode !== "bypassPermissions") {
        const dangerous = DANGEROUS_PATTERNS.find((p) => command.includes(p.pattern));
        if (dangerous) {
          return { block: true, reason: `⚠️ 危险命令已阻止: ${dangerous.description}` };
        }
      }
    }

    // ═══════════════════════════════════════════════════
    // 第二层：Plan 模式 — 工具级只读约束
    // ═══════════════════════════════════════════════════

    if (currentMode === "plan") {
      // 只读工具直接放行
      if (PLAN_TOOLS[tool]) return;

      // bash 仅放行安全白名单
      if (tool === "bash") {
        const command = typeof input.command === "string" ? input.command : "";
        const trimmed = command.trim();
        if (trimmed && SAFE_PLAN_BASH_PREFIXES.some((prefix) => trimmed.startsWith(prefix))) return;
        return { block: true, reason: "📋 Plan 模式：只允许只读工具和安全 bash 命令。请切换到 acceptEdits/bypassPermissions 模式执行。" };
      }

      // 其他工具全部阻止
      return { block: true, reason: "📋 Plan 模式：只读状态。请使用 read/search/find 工具探索代码，然后给出计划。" };
    }

    // ═══════════════════════════════════════════════════
    // 第三层：模式级审批逻辑
    // ═══════════════════════════════════════════════════

    // bypassPermissions 模式：全部放行（已经过第一层安全检查）
    if (currentMode === "bypassPermissions") return;

    // default 模式：所有工具都弹窗确认
    if (currentMode === "default") {
      if (isSessionAllowed(tool, input)) return;
      const ok = await ctx.ui.confirm(
        `允许执行 ${tool} 吗？`,
        ["允许", "拒绝"],
        "warning",
      );
      if (ok) {
        recordSessionApproval(tool, input);
        return;
      }
      return { block: true, reason: "用户拒绝" };
    }

    // acceptEdits 模式：write/edit 静默放行，执行类工具弹窗确认
    if (currentMode === "acceptEdits") {
      if (!EXEC_TOOLS[tool]) return;
      if (isSessionAllowed(tool, input)) return;
      const ok = await ctx.ui.confirm(
        `允许执行 ${tool} 吗？`,
        ["允许", "拒绝"],
        "warning",
      );
      if (ok) {
        recordSessionApproval(tool, input);
        return;
      }
      return { block: true, reason: "用户拒绝" };
    }

    return;
  });
}
