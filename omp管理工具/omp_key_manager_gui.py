"""
OMP API Key 管理器 - GUI 增强版
=================================
- Provider 下拉选择，自动读取 models.yml + models.json + auth_credentials
- 支持 API Key + BaseUrl 双字段
- 自动同步到 models.yml / models.json
- models.yml 不存在时自动创建
- 无 agent.db 时仍可正常操作配置文件
"""

import sqlite3, os, shutil, json
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

# ── 路径配置 ──────────────────────────────────────────────
HOME = os.path.expanduser("~")
DB_PATH = os.path.join(HOME, ".omp", "agent", "agent.db")
MODELS_YML = os.path.join(HOME, ".omp", "agent", "models.yml")
MODELS_JSON = os.path.join(HOME, ".omp", "agent", "models.json")
BACKUP_DIR = os.path.join(HOME, ".omp", "agent", "backups")
def _parse_yaml_simple(path=MODELS_YML):
    result = {}
    if not os.path.exists(path):
        return result
    try:
        with open(path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        current = None
        for line in lines:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            indent = len(line) - len(line.lstrip())
            if indent == 0 and stripped == "providers:":
                continue
            if indent == 2 and stripped.endswith(":"):
                current = stripped[:-1].strip()
                if current:
                    result[current] = {"source": "models.yml", "baseUrl": None, "apiKey": None, "models": []}
                continue
            if indent >= 4 and current:
                if stripped.startswith("baseUrl:"):
                    result[current]["baseUrl"] = stripped.split(":", 1)[1].strip().strip('"').strip("'")
                elif stripped.startswith("apiKey:"):
                    result[current]["apiKey"] = stripped.split(":", 1)[1].strip().strip('"').strip("'")
                elif stripped.startswith("- id:"):
                    result[current]["models"].append(stripped.split(":", 1)[1].strip())
        return result
    except Exception:
        return {}

def load_models_json():
    if not os.path.exists(MODELS_JSON):
        return {}
    try:
        with open(MODELS_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        result = {}
        for name, cfg in data.get("providers", {}).items():
            result[name] = {
                "source": "models.json",
                "baseUrl": cfg.get("baseUrl"),
                "apiKey": cfg.get("apiKey"),
                "models": [m.get("id", "") for m in cfg.get("models", []) if isinstance(m, dict)],
            }
        return result
    except Exception:
        return {}

def load_auth_credentials():
    if not os.path.exists(DB_PATH):
        return {}
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT provider, credential_type, data, disabled_cause FROM auth_credentials"
    )
    rows = cur.fetchall()
    conn.close()
    result = {}
    for provider, ctype, data_json, disabled in rows:
        if provider not in result:
            result[provider] = {
                "source": "auth_credentials", "baseUrl": None, "apiKey": None, "models": [],
                "disabled": disabled is not None,
            }
        try:
            parsed = json.loads(data_json)
        except json.JSONDecodeError:
            parsed = {"key": data_json}
        key_val = parsed.get("key", "")
        if ctype == "api_key":
            result[provider]["apiKey"] = key_val
        elif "://" in key_val:
            result[provider]["baseUrl"] = key_val
        elif not result[provider]["apiKey"]:
            result[provider]["apiKey"] = key_val
    return result

def merge_all_providers():
    merged = {}
    for loader in [_parse_yaml_simple, load_models_json, load_auth_credentials]:
        data = loader()
        for name, info in data.items():
            if name not in merged:
                merged[name] = dict(info)
            else:
                if not merged[name].get("baseUrl") and info.get("baseUrl"):
                    merged[name]["baseUrl"] = info["baseUrl"]
                if not merged[name].get("apiKey") and info.get("apiKey"):
                    merged[name]["apiKey"] = info["apiKey"]
                if not merged[name].get("models") and info.get("models"):
                    merged[name]["models"] = info["models"]
                if info.get("source") == "auth_credentials":
                    merged[name]["source"] = "auth_credentials"
                if info.get("disabled"):
                    merged[name]["disabled"] = True
    return merged

# ── 辅助函数 ──────────────────────────────────────────────

def mask_key(key, show=6):
    if not key:
        return "(未配置)"
    if len(key) <= show + 4:
        return key
    return key[:show] + "******" + key[-4:]

def backup_db():
    if not os.path.exists(DB_PATH):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bp = os.path.join(BACKUP_DIR, f"agent.db.{ts}.bak")
    shutil.copy2(DB_PATH, bp)
    return bp

def get_db():
    if not os.path.exists(DB_PATH):
        return None
    return sqlite3.connect(DB_PATH)

def fmt_ts(val):
    if isinstance(val, int):
        return datetime.fromtimestamp(val).strftime("%Y-%m-%d %H:%M")
    return str(val)

def read_yaml_lines(path):
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return f.readlines()

# ── models.yml / models.json 读写 ───────────────────────

def update_models_yml(provider_name, base_url, api_key):
    lines = read_yaml_lines(MODELS_YML)
    if lines is None:
        os.makedirs(os.path.dirname(MODELS_YML), exist_ok=True)
        with open(MODELS_YML, "w", encoding="utf-8") as f:
            f.write("providers:\n")
            f.write(f"  {provider_name}:\n")
            if base_url:
                f.write(f"    baseUrl: {base_url}\n")
            if api_key:
                f.write(f"    apiKey: {api_key}\n")
        return

    result = []
    in_provider = False
    fields_updated = set()
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        indent = len(line) - len(line.lstrip())

        if indent == 2 and stripped.endswith(":") and stripped[:-1].strip() == provider_name:
            in_provider = True
            result.append(line)
            i += 1
            continue

        if in_provider and indent >= 4:
            if indent == 2 and stripped.endswith(":"):
                if "baseUrl" not in fields_updated and base_url:
                    result.append(f"    baseUrl: {base_url}\n")
                if "apiKey" not in fields_updated and api_key:
                    result.append(f"    apiKey: {api_key}\n")
                in_provider = False
            elif stripped.startswith("baseUrl:"):
                if base_url:
                    result.append(f'{" " * indent}baseUrl: {base_url}\n')
                    fields_updated.add("baseUrl")
                i += 1
                continue
            elif stripped.startswith("apiKey:"):
                if api_key:
                    result.append(f'{" " * indent}apiKey: {api_key}\n')
                    fields_updated.add("apiKey")
                i += 1
                continue

        result.append(line)
        i += 1

    if in_provider:
        if "baseUrl" not in fields_updated and base_url:
            result.append("    baseUrl: " + base_url + "\n")
        if "apiKey" not in fields_updated and api_key:
            result.append("    apiKey: " + api_key + "\n")

    provider_in_file = any(f"{provider_name}:" in l for l in lines)
    if not provider_in_file and (base_url or api_key):
        result.append(f"  {provider_name}:\n")
        if base_url:
            result.append(f"    baseUrl: {base_url}\n")
        if api_key:
            result.append(f"    apiKey: {api_key}\n")

    with open(MODELS_YML, "w", encoding="utf-8") as f:
        f.writelines(result)

def sync_models_json():
    lines = read_yaml_lines(MODELS_YML)
    if lines is None:
        return
    providers = {}
    current = None
    for line in lines:
        stripped = line.strip()
        indent = len(line) - len(line.lstrip())
        if indent == 2 and stripped.endswith(":") and not stripped.startswith("-"):
            current = stripped[:-1].strip()
            providers[current] = {}
        elif indent >= 4 and current:
            if stripped.startswith("baseUrl:"):
                providers[current]["baseUrl"] = stripped.split(":", 1)[1].strip()
            elif stripped.startswith("apiKey:"):
                providers[current]["apiKey"] = stripped.split(":", 1)[1].strip()
    with open(MODELS_JSON, "w", encoding="utf-8") as f:
        json.dump({"providers": providers}, f, indent=2, ensure_ascii=False)

# ── Provider 编辑对话框 ──────────────────────────────

class ProviderEditDialog:
    def __init__(self, parent, existing_names):
        self.result = None
        self.win = tk.Toplevel(parent)
        self.win.title("添加 / 修改 Provider")
        self.win.geometry("540x370")
        self.win.resizable(False, False)
        self.win.transient(parent)
        self.win.grab_set()
        self.win.configure(bg="#f9f9f9")
        self.win.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 540) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 370) // 2
        self.win.geometry(f"+{max(0, x)}+{max(0, y)}")

        defaults = ["deepseek", "openai", "qwen", "vllm", "ollama", "local", "longi"]
        all_names = sorted(set(defaults + existing_names))

        tk.Label(self.win, text="Provider 名称:", font=("Microsoft YaHei", 10, "bold"),
                 bg="#f9f9f9").grid(row=0, column=0, sticky=tk.W, padx=16, pady=(14, 2))
        self.provider_var = tk.StringVar()
        combo = ttk.Combobox(self.win, textvariable=self.provider_var, width=36, state="normal")
        combo["values"] = all_names
        combo.grid(row=0, column=1, padx=6, pady=(14, 2), sticky=tk.W)
        combo.focus()
        tk.Label(self.win, text="(可输入新名称)", fg="#999", font=("Microsoft YaHei", 8),
                 bg="#f9f9f9").grid(row=0, column=2, sticky=tk.W)

        tk.Label(self.win, text="Base URL:", font=("Microsoft YaHei", 10, "bold"),
                 bg="#f9f9f9").grid(row=1, column=0, sticky=tk.W, padx=16, pady=(8, 2))
        self.url_var = tk.StringVar()
        tk.Entry(self.win, textvariable=self.url_var, width=40, font=("Consolas", 10)
                 ).grid(row=1, column=1, columnspan=2, padx=6, pady=(8, 2), sticky=tk.W)
        tk.Label(self.win, text="(vLLM/Ollama 自建必填)", fg="#999",
                 font=("Microsoft YaHei", 8), bg="#f9f9f9").grid(row=1, column=3, sticky=tk.W)

        tk.Label(self.win, text="API Key:", font=("Microsoft YaHei", 10, "bold"),
                 bg="#f9f9f9").grid(row=2, column=0, sticky=tk.W, padx=16, pady=(8, 2))
        self.key_var = tk.StringVar()
        tk.Entry(self.win, textvariable=self.key_var, width=40, font=("Consolas", 10), show="*"
                 ).grid(row=2, column=1, columnspan=2, padx=6, pady=(8, 2), sticky=tk.W)
        tk.Label(self.win, text="(云服务必填，自建可留空)", fg="#999",
                 font=("Microsoft YaHei", 8), bg="#f9f9f9").grid(row=2, column=3, sticky=tk.W)

        tip = tk.Frame(self.win, bg="#e8f4f8", bd=1, relief=tk.SOLID)
        tip.grid(row=3, column=0, columnspan=4, padx=16, pady=(12, 6), sticky=tk.W + tk.E)
        tk.Label(tip, text=(
            "  - DeepSeek/OpenAI  -> 只填 API Key，Base URL 留空\n"
            "  - vLLM/Ollama 自建 -> 只填 Base URL，API Key 留空\n"
            "  - 私有网关         -> 两者都填\n"
            "  - 修改后请重启 OMP 生效"
        ), font=("Microsoft YaHei", 9), bg="#e8f4f8", justify=tk.LEFT).pack(padx=8, pady=6)

        btns = tk.Frame(self.win, bg="#f9f9f9")
        btns.grid(row=4, column=0, columnspan=4, pady=(8, 8))
        tk.Button(btns, text="  确定  ", command=self._ok, bg="#27ae60", fg="white",
                  font=("Microsoft YaHei", 11, "bold"), bd=0, relief=tk.FLAT, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=6)
        tk.Button(btns, text="  取消  ", command=self.win.destroy, bg="#95a5a6", fg="white",
                  font=("Microsoft YaHei", 11, "bold"), bd=0, relief=tk.FLAT, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=6)

    def _ok(self):
        self.result = (
            self.provider_var.get().strip().lower(),
            self.url_var.get().strip(),
            self.key_var.get().strip(),
        )
        self.win.destroy()

# ── 主窗口 ──────────────────────────────────────────────

class KeyManagerApp:
    COLS = ("ID", "Provider", "Base URL", "API Key (脱敏)", "来源", "状态", "更新时间")
    WIDTHS = (40, 100, 240, 220, 90, 60, 120)

    def __init__(self, root):
        self.root = root
        self.root.title("OMP API Key 管理器")
        self.root.geometry("960x580")
        self.root.resizable(False, False)
        self.root.configure(bg="#f5f5f5")

        top = tk.Frame(root, bg="#2c3e50", height=50)
        top.pack(fill=tk.X)
        top.pack_propagate(False)
        tk.Label(top, text="OMP API Key 管理器  |  业务人员专用",
                 font=("Microsoft YaHei", 14, "bold"),
                 bg="#2c3e50", fg="#ecf0f1").pack(pady=8)

        tk.Label(root, text=DB_PATH, fg="#aaa", font=("Consolas", 8)
                 ).pack(anchor=tk.W, padx=14, pady=(2, 0))

        self._build_table()
        self._build_buttons()

        self.status_var = tk.StringVar(value="就绪")
        tk.Label(root, textvariable=self.status_var, bg="#f5f5f5",
                 fg="#666", font=("Microsoft YaHei", 9)
                 ).pack(side=tk.BOTTOM, fill=tk.X, padx=14, pady=(0, 6))

        self.refresh()

    def _build_table(self):
        frame = tk.Frame(self.root, bg="#f5f5f5")
        frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(4, 6))
        self.tree = ttk.Treeview(frame, columns=self.COLS, show="headings", height=16)
        for col, w in zip(self.COLS, self.WIDTHS):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor=tk.CENTER if col in ("ID", "状态") else tk.W)
        sb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Double-1>", self._on_double_click)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        conn = get_db()
        providers = merge_all_providers()

        if conn:
            cur = conn.execute(
                "SELECT id, provider, data, disabled_cause, updated_at "
                "FROM auth_credentials ORDER BY provider, id"
            )
            cred_rows = cur.fetchall()
            conn.close()
            seen = set()
            for rid, provider, data_json, disabled, updated in cred_rows:
                if rid in seen:
                    continue
                try:
                    parsed = json.loads(data_json)
                except json.JSONDecodeError:
                    parsed = {"key": data_json}
                key_val = parsed.get("key", "")
                pinfo = providers.get(provider, {})
                base_url = pinfo.get("baseUrl", "")
                if "://" in key_val and not base_url:
                    base_url = key_val
                    key_val = ""
                status = "已禁用" if disabled else "正常"
                source = pinfo.get("source", "auth_credentials")
                self.tree.insert("", tk.END, values=(
                    rid, provider, base_url or "-", mask_key(key_val), source, status, fmt_ts(updated)
                ))
                seen.add(rid)
            cred_names = {r[1] for r in cred_rows}
            for pname, pinfo in providers.items():
                if pname not in cred_names:
                    self.tree.insert("", tk.END, values=(
                        "-", pname, pinfo.get("baseUrl") or "-",
                        mask_key(pinfo.get("apiKey", "")), "models.yml", "正常", "-"
                    ))
        else:
            for pname, pinfo in providers.items():
                self.tree.insert("", tk.END, values=(
                    "-", pname, pinfo.get("baseUrl") or "-",
                    mask_key(pinfo.get("apiKey", "")), "models.yml", "正常", "-"
                ))

        self.status_var.set(f"共 {len(self.tree.get_children())} 个 Provider")

    def _build_buttons(self):
        bar = tk.Frame(self.root, bg="#f5f5f5")
        bar.pack(fill=tk.X, padx=14, pady=(0, 4))
        for text, cmd, color in [
            ("刷新列表", self.refresh, "#3498db"),
            ("添加 / 修改", self.add_update, "#27ae60"),
            ("删除", self.delete, "#e74c3c"),
            ("禁用", self.disable, "#f39c12"),
            ("启用", self.enable, "#1abc9c"),
        ]:
            tk.Button(bar, text=text, command=cmd, font=("Microsoft YaHei", 10, "bold"),
                      bg=color, fg="white", bd=0, relief=tk.FLAT,
                      activebackground=color, activeforeground="white",
                      cursor="hand2", padx=12, pady=6).pack(side=tk.LEFT, padx=3)

    def _selected_id(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先在表格中选择一行")
            return None
        return int(self.tree.item(sel[0])["values"][0])

    def _confirm(self, msg):
        return messagebox.askokcancel("确认", msg, icon="warning")

    def add_update(self):
        dlg = ProviderEditDialog(self.root, list(merge_all_providers().keys()))
        if not dlg.result:
            return
        provider, base_url, api_key = dlg.result
        if not provider:
            messagebox.showwarning("提示", "Provider 不能为空")
            return

        conn = get_db()
        now = int(datetime.now().timestamp())
        existing = None
        if conn:
            cur = conn.execute(
                "SELECT id FROM auth_credentials WHERE provider = ? AND credential_type = 'api_key'"
            )
            existing = cur.fetchone()

        if conn:
            backup_db()
            if existing:
                conn.execute(
                    "UPDATE auth_credentials SET data = ?, disabled_cause = NULL, updated_at = ? WHERE id = ?",
                    (json.dumps({"key": api_key}), now, existing[0])
                )
                conn.commit()
                action = "更新"
            elif api_key:
                conn.execute(
                    "INSERT INTO auth_credentials (provider, credential_type, data, created_at, updated_at) "
                    "VALUES (?, 'api_key', ?, ?, ?)",
                    (provider, json.dumps({"key": api_key}), now, now)
                )
                conn.commit()
                action = "添加"
            else:
                action = "仅写入配置文件"
            conn.close()
        else:
            action = "仅写入配置文件"

        update_models_yml(provider, base_url or None, api_key or None)
        sync_models_json()
        self.refresh()

        detail = []
        if base_url:
            detail.append(f"BaseUrl: {base_url}")
        if api_key:
            detail.append("API Key: 已设置")
        messagebox.showinfo("完成", f"Provider '{provider}' 已{action}\n" + "\n".join(detail))

    def delete(self):
        kid = self._selected_id()
        if kid is None:
            return
        conn = get_db()
        if not conn:
            return
        cur = conn.execute("SELECT provider FROM auth_credentials WHERE id = ?", (kid,))
        row = cur.fetchone()
        if not row:
            conn.close()
            return
        if not self._confirm(f"确定删除 Provider '{row[0]}' 的 Key？"):
            conn.close()
            return
        backup_db()
        conn.execute("DELETE FROM auth_credentials WHERE id = ?", (kid,))
        conn.commit()
        conn.close()
        self.refresh()
        messagebox.showinfo("完成", f"已删除 '{row[0]}'")

    def disable(self):
        kid = self._selected_id()
        if kid is None:
            return
        conn = get_db()
        if not conn:
            return
        cur = conn.execute("SELECT provider FROM auth_credentials WHERE id = ?", (kid,))
        row = cur.fetchone()
        conn.close()
        if not row:
            return
        if not self._confirm(f"确定禁用 '{row[0]}'？"):
            return
        conn = get_db()
        if not conn:
            return
        backup_db()
        conn.execute("UPDATE auth_credentials SET disabled_cause = '手动禁用', updated_at = ? WHERE id = ?",
                     (int(datetime.now().timestamp()), kid))
        conn.commit()
        conn.close()
        self.refresh()

    def enable(self):
        kid = self._selected_id()
        if kid is None:
            return
        conn = get_db()
        if not conn:
            return
        cur = conn.execute("SELECT provider, disabled_cause FROM auth_credentials WHERE id = ?", (kid,))
        row = cur.fetchone()
        conn.close()
        if not row:
            return
        if row[1] is None:
            messagebox.showinfo("提示", f"'{row[0]}' 已经是正常状态")
            return
        conn = get_db()
        if not conn:
            return
        conn.execute("UPDATE auth_credentials SET disabled_cause = NULL, updated_at = ? WHERE id = ?",
                     (int(datetime.now().timestamp()), kid))
        conn.commit()
        conn.close()
        self.refresh()
        messagebox.showinfo("完成", f"已启用 '{row[0]}'")

    def _on_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0])["values"]
        providers = merge_all_providers()
        pname = vals[1]
        extra = providers.get(pname, {})
        lines = [f"Provider: {vals[1]}", f"Base URL: {vals[2]}", f"API Key: {vals[3]}",
                 f"来源: {vals[4]}", f"状态: {vals[5]}"]
        if extra.get("models"):
            lines.append(f"\n可用模型: {', '.join(extra['models'])}")
        messagebox.showinfo("完整信息", "\n".join(lines))

if __name__ == "__main__":
    root = tk.Tk()
    KeyManagerApp(root)
    root.mainloop()
