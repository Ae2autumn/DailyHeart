# 设置控制台编码为UTF-8，确保中文正常显示
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# 更改控制台代码页为UTF-8 (65001)
chcp 65001 | Out-Null

# 加载Windows运行时程序集
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]

# 显示Toast通知的函数（包含错误捕获）
function Show-ToastNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$SubMessage,
        [string]$LogoPath,
        [switch]$IsError
    )
    
    try {
        # 根据是否有子消息选择模板
        $templateType = if ($SubMessage) { 
            [Windows.UI.Notifications.ToastTemplateType]::ToastText04 
        } else { 
            [Windows.UI.Notifications.ToastTemplateType]::ToastText02 
        }
        
        $toastXml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent($templateType)
        
        # 设置通知内容
        $toastXml.GetElementsByTagName("text").Item(0).AppendChild(
            $toastXml.CreateTextNode($Title)) | Out-Null
        $toastXml.GetElementsByTagName("text").Item(1).AppendChild(
            $toastXml.CreateTextNode($Message)) | Out-Null
        
        if ($SubMessage) {
            $toastXml.GetElementsByTagName("text").Item(2).AppendChild(
                $toastXml.CreateTextNode($SubMessage)) | Out-Null
        }
        
        # 添加Logo（错误通知不加Logo）
        if ($LogoPath -and (-not $IsError)) {
            $image = $toastXml.CreateElement("image")
            $image.SetAttribute("src", $LogoPath)
            $image.SetAttribute("placement", "appLogoOverride")
            $toastXml.GetElementsByTagName("binding").Item(0).AppendChild($image) | Out-Null
        }
        
        # 显示通知
        $toast = [Windows.UI.Notifications.ToastNotification]::new($toastXml)
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("PowerShell").Show($toast)
    }
    catch {
        # 如果Toast通知失败，回退到控制台输出
        Write-Host "Toast通知显示失败: $_" -ForegroundColor Red
        Write-Host "$Title", "$Message" -ForegroundColor Yellow
        if ($SubMessage) {
            Write-Host "$SubMessage" -ForegroundColor Yellow
        }
        return $false
    }
    return $true
}

# 主程序
try {
    # 读取配置
    $config = Get-Content "$PSScriptRoot./config.json" -ErrorAction Stop | ConvertFrom-Json
    $ApiUrl = $config.ApiUrl
    $Title = $config.Title
    $Signature = $config.Signature
    $LogoPath = $config.LogoPath
    $soundPath = $config.SoundPath

    # 检查VPN状态
    $vpn = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match "VPN" -and $_.Status -eq "Up" }

    if ($vpn) {
        # VPN警告通知
        $success = Show-ToastNotification -Title "VPN Was Enabled" `
            -Message "Please Turn Off VPN To Use" `
            -LogoPath $LogoPath
        
        if (-not $success) { exit 1 }
        exit 0
    }

    # 获取一言内容
    try {
        $response = Invoke-RestMethod -Uri $ApiUrl -ErrorAction Stop
        $message = $response.hitokoto
    }
    catch {
        # API请求失败，通过通知显示错误
        $success = Show-ToastNotification -Title "服务异常" `
            -Message "获取一言内容失败" `
            -SubMessage "$($_.Exception.Message)" `
            -IsError
        
        if (-not $success) { exit 1 }
        exit 0
    }

    # 显示一言通知
    $success = Show-ToastNotification -Title $Title `
        -Message $message `
        -SubMessage "Signature:`n$Signature" `
        -LogoPath $LogoPath
    
    if (-not $success) { exit 1 }
}
catch {
    # 主流程捕获的未处理异常
    $success = Show-ToastNotification -Title "系统错误" `
        -Message "脚本执行遇到问题" `
        -SubMessage "错误: $($_.Exception.Message)" `
        -IsError
    
    if (-not $success) { 
        Write-Host "主流程错误: $_" -ForegroundColor Red
        exit 1 
    }
}