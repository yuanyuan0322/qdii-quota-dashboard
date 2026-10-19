# -*- coding: utf-8 -*-
"""
QDII「海外额度宝」额度看板 —— Web 服务

功能：
  * 展示 QDII 基金单日申购额度排行（741 只份额）
  * 内置每日定时同步（默认 08:30），页面自动感知数据更新
  * 提供「立即同步」按钮，后台子进程执行 sync_data.py
  * 页面每 5 秒轮询状态，数据更新后自动刷新

启动：
  双击 start_server.bat   或   python app.py
访问：
  http://localhost:5001
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime, timedelta

from flask import Flask, jsonify, render_template

HERE = os.path.dirname(os.path.abspath(__file__))
WEB_JSON = os.path.join(HERE, "data", "quota_data.json")
SYNC_SCRIPT = os.path.join(HERE, "sync_data.py")
LOG_FILE = os.path.join(HERE, "sync.log")

PYTHON = sys.executable or r"C:\efunds\Python310\python.exe"
PORT = int(os.environ.get("QDII_PORT", "5001"))
SYNC_HOUR = int(os.environ.get("QDII_SYNC_HOUR", "8"))
SYNC_MINUTE = int(os.environ.get("QDII_SYNC_MINUTE", "30"))
STALE_HOURS = float(os.environ.get("QDII_STALE_HOURS", "20"))

app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False

_state = {
    "syncing": False,
    "progress": "",
    "last_result": "",
    "last_start": None,
    "last_finish": None,
    "log_tail": [],
}
_lock = threading.Lock()


# --------------------------------------------------------------------------- #
def load_data() -> dict:
    if not os.path.exists(WEB_JSON):
        return {"updated_at": None, "items": [], "stats": {}, "source": ""}
    try:
        with open(WEB_JSON, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        return {"updated_at": None, "items": [], "stats": {}, "error": str(e)}


def data_age_hours():
    d = load_data()
    ts = d.get("updated_at")
    if not ts:
        return None
    try:
        return (datetime.now() - datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")).total_seconds() / 3600
    except Exception:
        return None


def next_sync_time() -> str:
    now = datetime.now()
    t = now.replace(hour=SYNC_HOUR, minute=SYNC_MINUTE, second=0, microsecond=0)
    if now >= t:
        t += timedelta(days=1)
    return t.strftime("%Y-%m-%d %H:%M")


def append_log(line: str):
    with _lock:
        _state["log_tail"].append(line)
        _state["log_tail"] = _state["log_tail"][-40:]
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def run_sync(reason: str = "manual") -> bool:
    """在后台子进程里执行同步，返回是否已启动"""
    with _lock:
        if _state["syncing"]:
            return False
        _state["syncing"] = True
        _state["progress"] = "启动同步…"
        _state["last_start"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        _state["log_tail"] = []
    append_log(f"===== 同步开始（{reason}） {_state['last_start']} =====")

    def worker():
        rc = -1
        try:
            creation = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
            proc = subprocess.Popen(
                [PYTHON, "-u", SYNC_SCRIPT],
                cwd=HERE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="ignore",
                bufsize=1,
                creationflags=creation,
            )
            for line in proc.stdout:
                line = line.rstrip()
                if line:
                    append_log(line)
                    with _lock:
                        _state["progress"] = line
            rc = proc.wait()
        except Exception as e:
            append_log(f"同步异常：{e}")
        finally:
            with _lock:
                _state["syncing"] = False
                _state["last_result"] = "成功" if rc == 0 else f"失败(rc={rc})"
                _state["last_finish"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                _state["progress"] = "同步完成" if rc == 0 else _state["progress"]
            append_log(f"===== 同步结束 rc={rc} =====")

    threading.Thread(target=worker, name="sync-worker", daemon=True).start()
    return True


def scheduler_loop():
    """每 30 秒检查一次：到点且当天未同步过就触发"""
    time.sleep(20)  # 让服务先起来
    while True:
        try:
            d = load_data()
            last_date = (d.get("updated_at") or "")[:10]
            now = datetime.now()
            today = now.strftime("%Y-%m-%d")
            target = now.replace(hour=SYNC_HOUR, minute=SYNC_MINUTE, second=0, microsecond=0)
            if now >= target and last_date != today and not _state["syncing"]:
                append_log(f"定时任务触发（每日 {SYNC_HOUR:02d}:{SYNC_MINUTE:02d}）")
                run_sync("scheduled")
        except Exception as e:
            append_log(f"scheduler error: {e}")
        time.sleep(30)


def startup_check():
    """启动时若数据过期则自动补一次同步"""
    time.sleep(5)
    age = data_age_hours()
    if age is None:
        append_log("未发现看板数据，自动执行首次同步")
        run_sync("startup-no-data")
    elif age > STALE_HOURS:
        append_log(f"数据已 {age:.1f} 小时未更新，自动执行同步")
        run_sync("startup-stale")


# --------------------------------------------------------------------------- #
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/meta")
def api_meta():
    d = load_data()
    return jsonify({
        "updated_at": d.get("updated_at"),
        "count": len(d.get("items") or []),
        "stats": d.get("stats") or {},
        "source": d.get("source") or "",
        "next_sync": next_sync_time(),
        "interval_hint": f"每日 {SYNC_HOUR:02d}:{SYNC_MINUTE:02d} 自动同步",
        "syncing": _state["syncing"],
        "progress": _state["progress"],
        "last_result": _state["last_result"],
        "last_start": _state["last_start"],
        "last_finish": _state["last_finish"],
        "log_tail": _state["log_tail"][-12:],
        "age_hours": data_age_hours(),
    })


@app.route("/api/data")
def api_data():
    return jsonify(load_data())


@app.route("/api/sync", methods=["POST"])
def api_sync():
    started = run_sync("manual")
    return jsonify({"ok": started, "msg": "已开始同步" if started else "同步正在进行中"})


@app.route("/download/xlsx")
def download_xlsx():
    from flask import send_file
    p = os.path.join(HERE, "天天基金_海外额度宝_额度排行.xlsx")
    if not os.path.exists(p):
        return "Excel 尚未生成", 404
    return send_file(p, as_attachment=True,
                     download_name="天天基金_海外额度宝_额度排行.xlsx")


def main():
    print(f"QDII 额度看板已启动：http://localhost:{PORT}")
    print(f"每日自动同步时间：{SYNC_HOUR:02d}:{SYNC_MINUTE:02d}")
    threading.Thread(target=scheduler_loop, daemon=True).start()
    threading.Thread(target=startup_check, daemon=True).start()
    app.run(host="0.0.0.0", port=PORT, debug=False, threaded=True, use_reloader=False)


if __name__ == "__main__":
    main()
