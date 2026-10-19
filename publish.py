# -*- coding: utf-8 -*-
"""
把当天生成的静态看板快照发布到 GitHub Pages（docs/ 目录）

流程：
  1. QDII额度看板.html        -> docs/index.html      （GitHub Pages 入口）
  2. data/quota_data.json     -> docs/data/quota_data.json（可选，便于以后改 CDN 拉数）
  3. git add / commit / push

用法：
  python publish.py            # 发布（无变化则跳过 commit）
  python publish.py --no-push  # 只生成本地 docs，不提交推送

配合每日同步：在「立即同步」跑完后执行本脚本即可，或加进 Windows 计划任务。
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
SNAPSHOT = os.path.join(HERE, "QDII额度看板.html")
DATA_JSON = os.path.join(HERE, "data", "quota_data.json")
DOCS = os.path.join(HERE, "docs")


def run(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, cwd=HERE, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", check=check)


def main() -> int:
    if not os.path.exists(SNAPSHOT):
        print("[x] 未找到 QDII额度看板.html，请先执行同步")
        return 1

    os.makedirs(os.path.join(DOCS, "data"), exist_ok=True)
    shutil.copyfile(SNAPSHOT, os.path.join(DOCS, "index.html"))
    if os.path.exists(DATA_JSON):
        shutil.copyfile(DATA_JSON, os.path.join(DOCS, "data", "quota_data.json"))

    updated = ""
    n = 0
    if os.path.exists(DATA_JSON):
        try:
            with open(DATA_JSON, encoding="utf-8") as f:
                d = json.load(f)
            updated = d.get("updated_at", "")
            n = len(d.get("items", []))
        except Exception:
            pass
    stamp = updated or datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"[i] 快照已更新到 docs/index.html  (数据时间 {stamp}，{n} 只)")

    run(["git", "add", "-A"])
    status = run(["git", "status", "--porcelain"]).stdout.strip()
    if not status:
        print("[i] 无内容变化，跳过 commit")
        return 0

    msg = f"data: {stamp} · {n} 只 QDII"
    run(["git", "commit", "-m", msg])
    print(f"[√] 已提交：{msg}")

    if "--no-push" in sys.argv:
        print("[i] --no-push，未推送")
        return 0

    remotes = run(["git", "remote"]).stdout.strip()
    if not remotes:
        print("[!] 尚未配置远端仓库，先执行：")
        print("    git remote add origin https://github.com/<用户名>/<仓库>.git")
        print("    git push -u origin main")
        return 0

    if not run(["git", "config", "user.email"]).stdout.strip():
        print("[!] 未配置 git 身份，先执行：")
        print('    git config user.name "你的名字"')
        print('    git config user.email "你的邮箱"')
        return 0

    try:
        p = run(["git", "push"], check=False)
    except Exception as e:
        print(f"[x] push 异常：{e}")
        return 1
    print(p.stdout.strip() or p.stderr.strip())
    if p.returncode != 0:
        print("[x] 推送失败（常见原因：未登录 / token 过期 / 网络）")
        return 1
    print("[√] 已推送，GitHub Pages 约 1 分钟后生效")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
