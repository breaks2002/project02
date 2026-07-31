param(
    [string]$SapiPath = "$env:USERPROFILE\.claude\hooks\speak-notification-sapi.ps1"
)

$backupPath = "$SapiPath.original"

if (-not (Test-Path $SapiPath)) {
    Write-Host "错误：找不到语音文件" -ForegroundColor Red
    pause
    exit
}

if (-not (Test-Path $backupPath)) {
    Write-Host "首次运行，备份原文件..." -ForegroundColor Cyan
    Copy-Item -Path $SapiPath -Destination $backupPath -Force
}

$content = Get-Content $SapiPath -Raw -Encoding UTF8

if ($content -match "# VOICE: OFF") {
    Write-Host "当前状态：语音已关闭" -ForegroundColor Red
    Copy-Item -Path $backupPath -Destination $SapiPath -Force
    Write-Host "`n[已开启] 语音提示已开启" -ForegroundColor Green
} else {
    Write-Host "当前状态：语音已开启" -ForegroundColor Green
    
    $mutedContent = @"
# VOICE: OFF - 静音模式
param(
    [string]$Title = "Claude Code",
    [string]$Message = ""
)
# 所有语音代码已禁用
"@
    
    Set-Content -Path $SapiPath -Value $mutedContent -Encoding UTF8
    Write-Host "`n[已关闭] 语音提示已关闭" -ForegroundColor Red
}

[System.Media.SystemSounds]::Asterisk.Play()
Start-Sleep -Seconds 2