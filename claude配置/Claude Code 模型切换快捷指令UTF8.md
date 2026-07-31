# --- Claude Code 模型切换快捷指令 ---



# 强制PowerShell使用正确的编码读取自身
$global:EncodingFixApplied = $true

# 你的原有代码...



function use-db{
    $env:ANTHROPIC_AUTH_TOKEN = "你的key"
    $env:ANTHROPIC_BASE_URL = "https://ark.cn-beijing.volces.com/api/compatible"
    $env:ANTHROPIC_DEFAULT_OPUS_MODEL = "doubao-seed-code-preview-latest"
    $env:ANTHROPIC_DEFAULT_SONNET_MODEL = "doubao-seed-code-preview-latest"
    Write-Host "✅ Claude Code 环境已切换为 [豆包 Doubao]" -ForegroundColor Green
    Write-Host "🚀 正在启动 Claude Code..." -ForegroundColor Yellow
    
    # 启动 Claude Code
    claude --allow-dangerously-skip-permissions --permission-mode=plan @args
}

function use-ds {
    $env:ANTHROPIC_AUTH_TOKEN = "你的key"
    $env:ANTHROPIC_BASE_URL = "https://api.deepseek.com/anthropic"
    $env:ANTHROPIC_DEFAULT_OPUS_MODEL = "deepseek-chat"
    $env:ANTHROPIC_DEFAULT_SONNET_MODEL = "deepseek-chat"
    Write-Host "✅ Claude Code 环境已切换为 [DeepSeek]" -ForegroundColor Green
    Write-Host "🚀 正在启动 Claude Code..." -ForegroundColor Yellow
    
    # 启动 Claude Code
    claude --allow-dangerously-skip-permissions --permission-mode=plan @args
}

function use-cc {
    Write-Host "🚀 正在启动 Claude Code允许危险模式..." -ForegroundColor Yellow
    # 启动 Claude Code 开启危险模式
    claude --allow-dangerously-skip-permissions --permission-mode=plan @args
}