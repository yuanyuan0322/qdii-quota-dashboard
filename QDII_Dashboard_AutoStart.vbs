' ============================================================
' QDII Quota Dashboard - auto start + keepalive watchdog
' Runs start_server.bat every 5 minutes in a hidden window.
' start_server.bat checks port 5001 first, so the Flask server
' is started only when it is not already listening.
' Keep this file encoded as ANSI/GBK (Chinese path inside).
' ============================================================
Dim sh, bat
bat = "D:\opencode\project\oracle_xrisk\20261013_QDII¶î¶È\start_server.bat"
Set sh = CreateObject("WScript.Shell")
Do
    sh.Run "cmd /c """ & bat & """", 0, True
    WScript.Sleep 300000
Loop
