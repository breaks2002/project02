第一步：打开PowerShell配置文件

notepad $PROFILE

如果文件不存在，系统会提示创建。

第二步：添加模型切换函数

将以下配置添加到你的PowerShell配置文件中：
# --- Claude Code 模型切换快捷指令 ---

function use-glm {
    $env:ANTHROPIC_AUTH_TOKEN = "<your_glm_api_key>"
    $env:ANTHROPIC_BASE_URL = "https://open.bigmodel.cn/api/anthropic"
    $env:ANTHROPIC_DEFAULT_OPUS_MODEL = "glm-4.6"
    $env:ANTHROPIC_DEFAULT_SONNET_MODEL = "glm-4.6"
    $env:CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC = "1"
    $env:HTTP_PROXY = "http://127.0.0.1:7890"
    $env:HTTPS_PROXY = "http://127.0.0.1:7890"
    Write-Host "✅ Claude Code 环境已切换为 [智谱 GLM]" -ForegroundColor Green
    Write-Host "🚀 正在启动 Claude Code..." -ForegroundColor Yellow
    
    # 启动 Claude Code
    claude
}

function use-qwen {
    $env:ANTHROPIC_AUTH_TOKEN = "<your_qwen_api_key>"
    $env:ANTHROPIC_BASE_URL = "https://dashscope.aliyuncs.com/api/v2/apps/claude-code-proxy"
    $env:ANTHROPIC_DEFAULT_OPUS_MODEL = "qwen3-max"
    $env:ANTHROPIC_DEFAULT_SONNET_MODEL = "qwen3-max"
    $env:HTTP_PROXY = "http://127.0.0.1:7890"
    $env:HTTPS_PROXY = "http://127.0.0.1:7890"
    Write-Host "✅ Claude Code 环境已切换为 [通义千问 Qwen]" -ForegroundColor Green
    Write-Host "🚀 正在启动 Claude Code..." -ForegroundColor Yellow
    
    # 启动 Claude Code
    claude
}

function use-doubao {
    $env:ANTHROPIC_AUTH_TOKEN = "<your_doubao_api_key>"
    $env:ANTHROPIC_BASE_URL = "https://ark.cn-beijing.volces.com/api/compatible"
    $env:ANTHROPIC_DEFAULT_OPUS_MODEL = "doubao-seed-code-preview-latest"
    $env:ANTHROPIC_DEFAULT_SONNET_MODEL = "doubao-seed-code-preview-latest"
    $env:HTTP_PROXY = "http://127.0.0.1:7890"
    $env:HTTPS_PROXY = "http://127.0.0.1:7890"
    Write-Host "✅ Claude Code 环境已切换为 [豆包 Doubao]" -ForegroundColor Green
    Write-Host "🚀 正在启动 Claude Code..." -ForegroundColor Yellow
    
    # 启动 Claude Code
    claude
}

function use-ds {
    $env:ANTHROPIC_AUTH_TOKEN = "<your_deepseek_api_key>"
    $env:ANTHROPIC_BASE_URL = "https://api.deepseek.com/anthropic"
    $env:ANTHROPIC_DEFAULT_OPUS_MODEL = "deepseek-chat"
    $env:ANTHROPIC_DEFAULT_SONNET_MODEL = "deepseek-chat"
    $env:CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC = "1"
    $env:HTTP_PROXY = "http://127.0.0.1:7890"
    $env:HTTPS_PROXY = "http://127.0.0.1:7890"
    Write-Host "✅ Claude Code 环境已切换为 [DeepSeek]" -ForegroundColor Green
    Write-Host "🚀 正在启动 Claude Code..." -ForegroundColor Yellow
    
    # 启动 Claude Code
    claude
}

function use-ds-r1 {
    $env:ANTHROPIC_AUTH_TOKEN = "<your_deepseek_api_key>"
    $env:ANTHROPIC_BASE_URL = "https://api.deepseek.com/anthropic"
    $env:ANTHROPIC_DEFAULT_OPUS_MODEL = "deepseek-reasoner"
    $env:ANTHROPIC_DEFAULT_SONNET_MODEL = "deepseek-reasoner"
    $env:CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC = "1"
    $env:HTTP_PROXY = "http://127.0.0.1:7890"
    $env:HTTPS_PROXY = "http://127.0.0.1:7890"
    Write-Host "✅ Claude Code 环境已切换为 [DeepSeek-R1 推理模型]" -ForegroundColor Green
    Write-Host "🚀 正在启动 Claude Code..." -ForegroundColor Yellow
    
    # 启动 Claude Code
    claude
}


第三步：配置说明
关键环境变量：

 ANTHROPIC_AUTH_TOKEN: 各平台的API密钥

 ANTHROPIC_BASE_URL: 模型的API端点

 ANTHROPIC_DEFAULT_OPUS_MODEL: 使用的模型名称

 HTTP_PROXY/HTTPS_PROXY: 网络代理配置（根据实际情况调整）

代理配置说明：
 如果你不需要代理，可以删除这两行：

$env:HTTP_PROXY = "http://127.0.0.1:7890"
$env:HTTPS_PROXY = "http://127.0.0.1:7890"


使用方法
1. 重新加载配置
保存文件后，在PowerShell中执行：

powershell

. $PROFILE



集成命令

function cc {
    Write-Host "🚀 正在启动 Claude Code允许危险模式..." -ForegroundColor Yellow
    # 启动 Claude Code 开启危险模式
    claude --allow-dangerously-skip-permissions --permission-mode=plan
}