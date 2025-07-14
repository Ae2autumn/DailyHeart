$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $scriptDir

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $scriptDir

# 读取配置文件
$config = Get-Content "$PSScriptRoot./config.json" | ConvertFrom-Json
$ApiUrl = $config.ApiUrl
$Title = $config.Title
$Signature = $config.Signature
$LogoPath = $config.LogoPath
$soundPath = $config.SoundPath

# 检查VPN状态
$vpn = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match "VPN" -and $_.Status -eq "Up" }

try {
    if ($vpn) {
        # 如果VPN开启，发送警告通知
        $vpnMessage = "检测到VPN已开启"
        $vpnWarning = "请关闭VPN后再使用此功能"
        New-BurntToastNotification -Text $vpnMessage, $vpnWarning -AppLogo $LogoPath -Sound $config.Sound
        exit 0
    }
    else {
        # 正常获取一言内容
        $response = Invoke-RestMethod -Uri $ApiUrl -ErrorAction Stop
        $message = $response.hitokoto #  + " -" + $response.from(添加源)
        New-BurntToastNotification -Text $Title, $message, Signature:`n$Signature -AppLogo $LogoPath
    }
}
catch {
    Write-Host "操作失败: $_" -ForegroundColor Red
    exit 1
}