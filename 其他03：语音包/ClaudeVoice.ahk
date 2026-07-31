; Claude 语音切换快捷键 - AutoHotkey v2 版本
; Alt + Shift + V 切换开关
!+v::
{
    Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"C:\Users\shifei2\.claude\hooks\toggle-claude-voice.ps1`""
}

; Ctrl + Alt + V 显示当前状态（改用弹窗）
^!v::
{
    Content := FileRead("C:\Users\shifei2\.claude\hooks\speak-notification-sapi.ps1")
    if InStr(Content, "# VOICE: OFF")
        MsgBox "🔇 Claude 语音当前已关闭", "语音状态", "64 T2"
    else
        MsgBox "🔊 Claude 语音当前已开启", "语音状态", "64 T2"
}