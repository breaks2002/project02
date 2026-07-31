param(
    [string]$Title = "",
    [string]$Category = "SessionStart"
)

# 自动获取当前用户的主目录
$userProfile = $env:USERPROFILE
# 或者使用：$home 变量也可以

# 为不同类别定义消息数组
$messages = @{
    "SessionStart" = @(
        "会话已开启，随时待命",
        "我准备好了，有什么需要帮忙的",
        "早上好下午好晚上好，开始工作吧",
        "您的AI助理已就绪，飞哥，请指示",
        "嗨，飞哥，我在这里，有什么问题"
    )
    
    "SessionEnd" = @(
        "会话结束，下次再见",
        "今天就到这里，拜拜",
        "会话已结束，期待下次见面",
        "工作完成，关闭会话",
        "再见飞哥，有需要随时找我"
    )
    
    "UserPromptSubmit" = @(
        "好的，我来处理",
        "收到请求，正在思考",
        "让我看看这个问题",
        "好的，请稍等",
        "收到，马上处理"
    )
    
    "PreCompact" = @(
        "信息有点多，我梳理一下再继续",
        "整理一下思路，请稍等",
        "让我理清这些信息",
        "正在梳理上下文",
        "信息整合中，马上好"
    )
    
    "PreToolUse" = @(
        "准备执行操作了",
        "马上调用工具处理",
        "准备使用工具",
        "开始执行任务",
        "正在准备工具"
    )
    
    "PostToolUse" = @(
        "操作已完成",
        "工具执行完毕",
        "任务完成",
        "操作成功",
        "搞定"
    )
    
    "PostToolUseFailure" = @(
        "操作出了点小问题，我看看怎么回事",
        "工具执行遇到问题，正在排查",
        "出错了，让我想想",
        "操作失败，换个方式试试",
        "遇到点麻烦，稍等"
    )
    
    "Stop" = @(
        "任务已完成，请查收",
        "所有任务完成",
        "搞定，看看结果",
        "任务结束，结果在此",
        "完成了，请检查"
    )
    
    "Notification-permission_prompt" = @(
        "需要您授权才能继续",
        "飞哥，请授权操作",
        "需要您的许可",
        "授权后才能继续",
        "飞哥，请确认权限"
    )
    
    "Notification-idle_prompt" = @(
        "等待您的输入",
        "我在等您的指令",
        "请告诉我下一步",
        "需要您的指示",
        "等您说话呢"
    )
    
    "SubagentStart" = @(
        "我启动个小助手专门处理这个任务",
        "开启子任务处理",
        "派个小助手处理",
        "启动辅助代理",
        "让助手处理这个"
    )
    
    "SubagentStop" = @(
        "小助手任务完成，结果已返回",
        "助手任务结束",
        "子任务完成",
        "助手返回结果",
        "小任务搞定"
    )
    
    "TaskCompleted" = @(
        "大功告成",
        "任务圆满完成",
        "完美收工",
        "任务成功",
        "搞定收工"
    )
    
    "PermissionRequest" = @(
        "需要您授权才能继续",
        "飞哥，请授权操作",
        "飞哥，需要您的许可",
        "授权后才能继续",
        "飞哥，请确认权限"
    )
}

# 获取随机消息
function Get-RandomMessage {
    param([string]$category)
    
    if ($messages.ContainsKey($category)) {
        $messageList = $messages[$category]
        $randomIndex = Get-Random -Minimum 0 -Maximum $messageList.Count
        return $messageList[$randomIndex]
    } else {
        return "Claude处理中"
    }
}

# 获取随机消息
$message = Get-RandomMessage -category $Category

# 动态构建原始脚本路径
$originalScript = Join-Path $userProfile ".claude\hooks\speak-notification-sapi.ps1"

# 调用原来的语音脚本
& $originalScript -Title $Title -Message $message