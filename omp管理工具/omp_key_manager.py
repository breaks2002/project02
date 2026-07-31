"""
OMP API Key 管理器
==================
可视化管理 OMP 凭据库中的 API Key（agent.db -> auth_credentials 表）
支持查看、修改、删除操作，无需手动敲 SQL。
"""

import sqlite3
import os
import shutil
import json
from datetime import datetime

# ── 路径配置 ──────────────────────────────────────────────
DB_PATH = os.path.expanduser(r"~\.omp\agent\agent.db")
BACKUP_DIR = os.path.expanduser(r"~\.omp\agent\backups")

# ── 辅助函数 ──────────────────────────────────────────────

def mask_key(key: str, show: int = 6) -> str:
    """脱敏显示 key，只保留前 show 位和末 4 位"""
    key = key.strip()
    if len(key) <= show + 4:
        return key
    return key[:show] + "******" + key[-4:]


def extract_key(data_json: str) -> str:
    """从 JSON 中提取 key 字符串"""
    try:
        parsed = json.loads(data_json)
        return parsed.get("key", data_json)
    except json.JSONDecodeError:
        return data_json


def backup_db():
    """修改前自动备份"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BACKUP_DIR, f"agent.db.{ts}.bak")
    shutil.copy2(DB_PATH, backup_path)
    print(f"  [OK] 已自动备份到: {backup_path}\n")
    return backup_path


def get_connection():
    """获取数据库连接"""
    if not os.path.exists(DB_PATH):
        print(f"\n  [!!] 未找到数据库文件: {DB_PATH}")
        print(f"       请确认 OMP 已至少运行过一次。\n")
        return None
    return sqlite3.connect(DB_PATH)


# ── 核心功能 ──────────────────────────────────────────────

def list_keys(conn):
    """列出所有凭据"""
    cur = conn.execute(
        "SELECT id, provider, credential_type, data, disabled_cause, created_at, updated_at "
        "FROM auth_credentials ORDER BY provider, id"
    )
    rows = cur.fetchall()

    if not rows:
        print("\n  [..] 没有找到任何 API Key 记录。\n")
        return

    print()
    print("=" * 70)
    print(f"  OMP API Key 列表（共 {len(rows)} 条）")
    print("=" * 70)

    for row in rows:
        id_, provider, ctype, data_json, disabled, created, updated = row
        key_raw = extract_key(data_json)
        status = "[DISABLED]" if disabled else "[ACTIVE]"

        print(f"\n  ID: {id_}  |  Provider: {provider}  |  {status}")
        print(f"       Key: {mask_key(key_raw)}")
        ts_c = (datetime.fromtimestamp(created).strftime('%Y-%m-%d %H:%M')
                if isinstance(created, int) else created)
        ts_u = (datetime.fromtimestamp(updated).strftime('%Y-%m-%d %H:%M')
                if isinstance(updated, int) else updated)
        print(f"       创建: {ts_c}  |  更新: {ts_u}")

    print("\n" + "=" * 70 + "\n")


def add_or_update_key(conn):
    """添加或更新 API Key"""
    provider = input("  输入 Provider 名称（如 deepseek、vllm、openai）: ").strip().lower()
    if not provider:
        print("  [!!] Provider 名称不能为空\n")
        return

    api_key = input("  输入 API Key: ").strip()
    if not api_key:
        print("  [!!] API Key 不能为空\n")
        return

    cur = conn.execute(
        "SELECT id, data FROM auth_credentials WHERE provider = ? AND credential_type = 'api_key'",
        (provider,)
    )
    existing = cur.fetchone()
    now = int(datetime.now().timestamp())
    data_str = json.dumps({"key": api_key})

    if existing:
        print(f"\n  [!!] Provider '{provider}' 已存在（ID: {existing[0]}）")
        print(f"       当前 Key: {mask_key(existing[1])}")
        confirm = input(f"       是否覆盖为: {mask_key(data_str)}？(y/N): ").strip().lower()
        if confirm != "y":
            print("  [..] 已取消\n")
            return

        backup_db()
        conn.execute(
            "UPDATE auth_credentials SET data = ?, disabled_cause = NULL, updated_at = ? WHERE id = ?",
            (data_str, now, existing[0])
        )
        print(f"  [OK] Provider '{provider}' 的 Key 已更新\n")
    else:
        backup_db()
        conn.execute(
            "INSERT INTO auth_credentials (provider, credential_type, data, created_at, updated_at) "
            "VALUES (?, 'api_key', ?, ?, ?)",
            (provider, data_str, now, now)
        )
        print(f"  [OK] Provider '{provider}' 的 Key 已添加\n")

    conn.commit()


def delete_key(conn):
    """删除一条凭据"""
    list_keys(conn)

    try:
        key_id = input("  输入要删除的 ID: ").strip()
        if not key_id:
            return
        key_id = int(key_id)
    except ValueError:
        print("  [!!] 无效的 ID\n")
        return

    cur = conn.execute(
        "SELECT id, provider, data FROM auth_credentials WHERE id = ?", (key_id,)
    )
    row = cur.fetchone()
    if not row:
        print(f"  [!!] ID {key_id} 不存在\n")
        return

    print(f"\n  [!!] 即将删除:")
    print(f"       Provider: {row[1]}")
    print(f"       Key: {mask_key(extract_key(row[2]))}")

    confirm = input("  确认删除？此操作不可撤销！(y/N): ").strip().lower()
    if confirm != "y":
        print("  [..] 已取消\n")
        return

    backup_db()
    conn.execute("DELETE FROM auth_credentials WHERE id = ?", (key_id,))
    conn.commit()
    print(f"  [OK] 已删除 Provider '{row[1]}' 的 Key\n")


def disable_key(conn):
    """禁用一个 Key 而不是删除"""
    list_keys(conn)

    try:
        key_id = input("  输入要禁用的 ID: ").strip()
        if not key_id:
            return
        key_id = int(key_id)
    except ValueError:
        print("  [!!] 无效的 ID\n")
        return

    cur = conn.execute(
        "SELECT id, provider FROM auth_credentials WHERE id = ?", (key_id,)
    )
    row = cur.fetchone()
    if not row:
        print(f"  [!!] ID {key_id} 不存在\n")
        return

    reason = input(f"  禁用原因（可选，留空直接禁用）: ").strip() or None
    backup_db()
    conn.execute(
        "UPDATE auth_credentials SET disabled_cause = ?, updated_at = ? WHERE id = ?",
        (reason, int(datetime.now().timestamp()), key_id)
    )
    conn.commit()
    print(f"  [OK] 已禁用 Provider '{row[1]}' (ID: {key_id})\n")


def enable_key(conn):
    """重新启用一个被禁用的 Key"""
    cur = conn.execute(
        "SELECT id, provider, data FROM auth_credentials WHERE disabled_cause IS NOT NULL"
    )
    rows = cur.fetchall()

    if not rows:
        print("\n  [..] 没有已禁用的 Key。\n")
        return

    print("\n  已禁用的 Key:")
    for row in rows:
        print(f"    [{row[0]}] {row[1]} - {mask_key(extract_key(row[2]))}")

    try:
        key_id = input("\n  输入要启用的 ID: ").strip()
        if not key_id:
            return
        key_id = int(key_id)
    except ValueError:
        print("  [!!] 无效的 ID\n")
        return

    conn.execute(
        "UPDATE auth_credentials SET disabled_cause = NULL, updated_at = ? WHERE id = ?",
        (int(datetime.now().timestamp()), key_id)
    )
    conn.commit()
    print(f"  [OK] 已重新启用\n")


# ── 主菜单 ──────────────────────────────────────────────

def main():
    print()
    print("  ============================================")
    print("     OMP API Key 管理器")
    print("     配置文件: agent.db -> auth_credentials")
    print("  ============================================")
    print()

    conn = get_connection()
    if not conn:
        input("  按 Enter 退出...")
        return

    while True:
        print("  +------------------------------------------+")
        print("  |  1. 查看所有 API Key                      |")
        print("  |  2. 添加 / 修改 API Key                   |")
        print("  |  3. 删除 API Key                          |")
        print("  |  4. 禁用 API Key（保留记录）              |")
        print("  |  5. 启用已禁用的 Key                      |")
        print("  |  0. 退出                                  |")
        print("  +------------------------------------------+")
        print()

        choice = input("  >> 请输入选项 [0-5]: ").strip()

        if choice == "1":
            list_keys(conn)
        elif choice == "2":
            add_or_update_key(conn)
        elif choice == "3":
            delete_key(conn)
        elif choice == "4":
            disable_key(conn)
        elif choice == "5":
            enable_key(conn)
        elif choice == "0":
            print("\n  再见！\n")
            break
        else:
            print("  [!!] 无效选项，请重新输入\n")

    conn.close()


if __name__ == "__main__":
    main()
