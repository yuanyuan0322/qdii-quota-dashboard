# -*- coding: utf-8 -*-
"""生成 QDII 看板的开机自启 + 保活 VBS（ANSI/GBK 编码，中文路径安全）"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
BAT = os.path.join(HERE, "start_server.bat")

VBS = '''\' ============================================================
\' QDII Quota Dashboard - auto start + keepalive watchdog
\' Runs start_server.bat every 5 minutes in a hidden window.
\' start_server.bat checks port 5001 first, so the Flask server
\' is started only when it is not already listening.
\' Keep this file encoded as ANSI/GBK (Chinese path inside).
\' ============================================================
Dim sh, bat
bat = "{bat}"
Set sh = CreateObject("WScript.Shell")
Do
    sh.Run "cmd /c """ & bat & """", 0, True
    WScript.Sleep 300000
Loop
'''.format(bat=BAT)

MASTER = os.path.join(HERE, "QDII_Dashboard_AutoStart.vbs")
STARTUP = os.path.join(os.environ["APPDATA"],
                       r"Microsoft\Windows\Start Menu\Programs\Startup",
                       "QDII_Dashboard_AutoStart.vbs")

data = VBS.replace("\n", "\r\n").encode("gbk")
with open(MASTER, "wb") as f:
    f.write(data)
print("master :", MASTER, len(data), "bytes")

with open(STARTUP, "wb") as f:
    f.write(data)
print("startup:", STARTUP, len(data), "bytes")

# 自检：回读确认中文路径无损
back = open(STARTUP, "rb").read().decode("gbk")
assert BAT in back, "路径回读不一致"
print("verify : OK ->", BAT)
