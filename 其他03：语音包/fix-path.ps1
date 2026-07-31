# interactive-fix.ps1

$settingsFile = "$env:USERPROFILE\.claude\settings.json"

if (-not (Test-Path $settingsFile)) {
    Write-Host "? 找不到文件: $settingsFile" -ForegroundColor Red
    exit 1
}

$oldUser = Read-Host "请输入原用户名 (例如: shifei2)"
$newUser = $env:USERNAME

Write-Host "将把 '$oldUser' 替换为 '$newUser'" -ForegroundColor Yellow

$content = Get-Content $settingsFile -Raw
$newContent = $content -replace "C:/Users/$oldUser/", "C:/Users/$newUser/"

# 显示几行替换结果
Write-Host "`n替换后的路径示例:" -ForegroundColor Green
$newContent -split "`n" | Select-String "C:/Users/" | Select-Object -First 3 | ForEach-Object { Write-Host $_ -ForegroundColor Cyan }

$confirm = Read-Host "`n确认保存? (y/n)"
if ($confirm -eq 'y') {
    Copy-Item $settingsFile "$settingsFile.backup" -Force
    $newContent | Set-Content $settingsFile
    Write-Host "? 已保存！" -ForegroundColor Green
} else {
    Write-Host "已取消" -ForegroundColor Red
}

Read-Host "按回车键退出"