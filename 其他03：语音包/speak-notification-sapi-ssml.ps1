param(
    [string]$Title = "Claude Code",
    [string]$Message = ""
)

# 语音提示 - 使用同步播放
try {
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    
    # 查找并选择Xiaoxiao语音
    $voices = $synth.GetInstalledVoices() | Where-Object { $_.Enabled }
    $xiaoxiaoVoice = $voices | Where-Object { 
        $_.VoiceInfo.Name -like "*Xiaoxiao*" -or 
        $_.VoiceInfo.Name -like "*晓晓*" 
    } | Select-Object -First 1
    
    if ($xiaoxiaoVoice) {
        $synth.SelectVoice($xiaoxiaoVoice.VoiceInfo.Name)
    }
    
    # 设置正常语速
    $synth.Rate = 0
    
    # 构建 SSML 格式（带 version 属性）
    $ssml = @"
<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="zh-CN">
    $Title $Message
</speak>
"@
    
    # 使用 SSML 播放
    $synth.SpeakSsml($ssml)
    
} catch {
    # 忽略语音错误
}