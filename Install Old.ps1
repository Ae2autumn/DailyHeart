# 设置控制台编码为 UTF-8（保持原提示不变）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

# 检查管理员权限（保持原提示不变）
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "请以管理员权限运行此脚本。" -ForegroundColor Red
    Start-Sleep -Seconds 3
    Exit 1
}

# 统一的重启提示函数（保持原提示文本不变）
function Invoke-RestartPrompt {
    param([string]$operation)
    $restartChoice = Read-Host "$operation 完成，是否立即重启系统? (Y/N)"
    if ($restartChoice -eq "y" -or $restartChoice -eq "Y") {
        Write-Host "系统将在 5 秒后重启..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
        Restart-Computer -Force
    } else {
        Write-Host "请手动重启系统以完成所有更改" -ForegroundColor Yellow
    }
}

# 主菜单（保持原提示不变）
$userInput = Read-Host "确定要安装 DailyHeart 吗?`nI(nstall)/U(ninstall)/R(estore Language)"

# 安装功能（保持原提示不变）
if ($userInput -eq "i" -or $userInput -eq "I") {
    $scriptDir = $PSScriptRoot
    $sourceFolder = "$scriptDir\DailyHeart"
    $destinationPath = "C:\Program Files (x86)\"
    
    if (Test-Path $sourceFolder -PathType Container) {
        try {
            Move-Item -Path $sourceFolder -Destination $destinationPath -Force
            Write-Host "已成功移动安装文件夹" -ForegroundColor Green
            
            # 添加计划任务（保持原提示风格）
            $taskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-File `"$destinationPath\DailyHeart\main.ps1`""
            $taskTrigger = New-ScheduledTaskTrigger -Daily -At 9am
            Register-ScheduledTask -TaskName "DailyHeart" -Action $taskAction -Trigger $taskTrigger -Force
            Write-Host "已创建计划任务" -ForegroundColor Green
            
            # 设置控制台颜色（保持原警告文本不变）
            $host.ui.RawUI.BackgroundColor = "Black"
            $host.ui.RawUI.ForegroundColor = "Red"
            Write-Host "警告：启用全局 UTF-8 支持可能会导致某些程序出现异常！" -ForegroundColor Red
            Write-Host "如果不确定后果或无法恢复系统，请不要继续！" -ForegroundColor Red
            
            $confirmInput = Read-Host "Y(es)/N(o)"
            if ($confirmInput -eq "y" -or $confirmInput -eq "Y") {
                try {
                    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage"
                    Set-ItemProperty -Path $registryPath -Name "ACP" -Value "65001"
                    Set-ItemProperty -Path $registryPath -Name "OEMCP" -Value "65001"
                    Set-ItemProperty -Path $registryPath -Name "MACCP" -Value "65001"
                    Set-ItemProperty -Path $registryPath -Name "ANSI" -Value "65001"
                    Write-Host "已成功启用 UTF-8 支持！" -ForegroundColor Green
                } catch {
                    Write-Host "修改注册表时出错：$_" -ForegroundColor Red
                }
            }else {
                Write-Host "操作已取消。" -ForegroundColor Yellow
                Start-Sleep -Seconds 3
                Exit 0
            }
            
            Invoke-RestartPrompt -operation "安装"
        } catch {
            Write-Host "安装过程中出错: $_" -ForegroundColor Red
            Start-Sleep -Seconds 3
            Exit 1
        }
    } else {
        Write-Host "请检查完整性" -ForegroundColor Red
        Start-Sleep -Seconds 3
        Exit 1
    }
}
# 卸载功能（保持原提示不变）
elseif ($userInput -eq "u" -or $userInput -eq "U") {
    $taskName = "DailyHeart"
    $installPath = "C:\Program Files (x86)\DailyHeart"

    # 1. 删除计划任务（保持原提示不变）
    try {
        $taskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($taskExists) {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            Write-Host "已删除计划任务: $taskName" -ForegroundColor Green
        } else {
            Write-Host "计划任务不存在: $taskName" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "删除计划任务失败: $_" -ForegroundColor Red
    }

    # 2. 删除安装目录（保持原提示不变）
    if (Test-Path $installPath) {
        try {
            Remove-Item -Path $installPath -Recurse -Force
            Write-Host "已删除安装目录: $installPath" -ForegroundColor Green
        } catch {
            Write-Host "删除安装目录失败: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "您真的安装了吗？" -ForegroundColor Yellow
    }

    # 3. 自动恢复语言设置（保持原提示不变）
    try {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "ACP" -Value "936"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "OEMCP" -Value "936"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "MACCP" -Value "936"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "ANSI" -Value "936"
        Write-Host "已恢复默认语言设置(936 zh_CN)" -ForegroundColor Green
    } catch {
        Write-Host "恢复语言设置失败: $_" -ForegroundColor Red
    }

    Invoke-RestartPrompt -operation "卸载"
}
# 恢复语言功能（保持原提示不变）
elseif ($userInput -eq "r" -or $userInput -eq "R") {
    Write-Host "正在恢复默认语言设置(936 zh_CN)..." -ForegroundColor Yellow
    try {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "ACP" -Value "936"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "OEMCP" -Value "936"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "MACCP" -Value "936"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "ANSI" -Value "936"
        Write-Host "已恢复默认语言设置" -ForegroundColor Green
        Invoke-RestartPrompt -operation "语言恢复"
    } catch {
        Write-Host "恢复语言设置失败: $_" -ForegroundColor Red
    }
}
else {
    Write-Host "操作已取消。" -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    Exit 0
}