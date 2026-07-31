;==========================================================
; 豆包语音输入助手 - AutoHotKey v2 脚本
; 功能：按住右Alt键触发语音输入，松开自动插入，保护剪贴板
;==========================================================

#Requires AutoHotkey v2.0
#SingleInstance Force

; ==================== 配置区域 ====================
; 豆包语音输入快捷键
DOUBAO_HOTKEY := "^d"  ; Ctrl+D

; 松开后等待识别完成的时间（毫秒）
; 说话结束后豆包需要时间完成识别，建议1500-3000
WAIT_AFTER_RELEASE := 1500

; 是否显示提示
SHOW_TIPS := true
; =================================================

; 全局变量
global clipboardBackup := ""
global isVoiceActive := false
global keyHoldStart := 0

; 创建托盘菜单
A_TrayMenu.Delete()
A_TrayMenu.Add("增加等待时间 (+500ms)", IncreaseWait)
A_TrayMenu.Add("减少等待时间 (-500ms)", DecreaseWait)
A_TrayMenu.Add()
A_TrayMenu.Add("关于", ShowAbout)
A_TrayMenu.Add("退出", DoExit)
TraySetIcon("shell32.dll", 169)

if SHOW_TIPS
    TrayTip("豆包语音助手", "按住右Alt说话，松开自动插入`n当前等待: " . WAIT_AFTER_RELEASE . "ms", 1)

; ==================== 热键绑定 ====================

*RAlt::
{
    global clipboardBackup, isVoiceActive, keyHoldStart

    if isVoiceActive
        return

    isVoiceActive := true
    keyHoldStart := A_TickCount

    clipboardBackup := ClipboardAll()

    Send(DOUBAO_HOTKEY)

    if SHOW_TIPS
        ToolTip("正在录音...")
}

*RAlt Up::
{
    global isVoiceActive, keyHoldStart, WAIT_AFTER_RELEASE

    if !isVoiceActive
        return

    holdTime := A_TickCount - keyHoldStart

    ; 根据说话时长动态调整等待
    waitTime := WAIT_AFTER_RELEASE
    if holdTime > 5000
        waitTime += 500
    if holdTime > 10000
        waitTime += 500

    if SHOW_TIPS
        ToolTip("等待识别... (" . waitTime . "ms)")

    Sleep(waitTime)

    ; 发送回车到豆包弹窗
    Send("{Enter}")

    ToolTip()

    SetTimer(RestoreClipboard, -500)

    isVoiceActive := false
}

; ==================== 函数 ====================

RestoreClipboard()
{
    global clipboardBackup
    if clipboardBackup != ""
    {
        A_Clipboard := clipboardBackup
        clipboardBackup := ""
    }
}

IncreaseWait(*)
{
    global WAIT_AFTER_RELEASE
    WAIT_AFTER_RELEASE += 500
    TrayTip("等待时间已调整", "当前: " . WAIT_AFTER_RELEASE . "ms", 1)
}

DecreaseWait(*)
{
    global WAIT_AFTER_RELEASE
    if WAIT_AFTER_RELEASE > 500
        WAIT_AFTER_RELEASE -= 500
    TrayTip("等待时间已调整", "当前: " . WAIT_AFTER_RELEASE . "ms", 1)
}

ShowAbout(*)
{
    global WAIT_AFTER_RELEASE
    MsgBox("豆包语音输入助手`n`n按住右Alt说话，松开自动插入`n当前等待时间: " . WAIT_AFTER_RELEASE . "ms`n`n右键托盘可调整等待时间", "关于", "Iconi")
}

DoExit(*)
{
    ExitApp()
}