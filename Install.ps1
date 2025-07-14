# 设置控制台编码为 UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

# 检查管理员权限
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "请以管理员权限运行此脚本。" -ForegroundColor Red
    Start-Sleep -Seconds 3
    Exit 1
}

# 统一的重启提示函数（增加手动重启后的延迟）
function Invoke-RestartPrompt {
    param([string]$operation)
    do {
        $restartChoice = Read-Host "$operation 完成，是否立即重启系统? (Y/N)"
        switch ($restartChoice.ToUpper()) {
            "Y" {
                Write-Host "系统将在 5 秒后重启..." -ForegroundColor Yellow
                Start-Sleep -Seconds 5
                Restart-Computer -Force
                break
            }
            "N" {
                Write-Host "请手动重启系统以完成所有更改" -ForegroundColor Yellow
                Start-Sleep -Seconds 5
                Exit 0
            }
            default {
                Write-Host "无效，请重新输入" -ForegroundColor Red
            }
        }
    } while ($true)
}

# 卸载函数
function Uninstall-DailyHeart {
    param(
        [bool]$Partial = $false
    )

    $taskName = "DailyHeart"
    $installPath = "C:\Program Files (x86)\DailyHeart"

    Uninstall-Module -Name BurntToast -Force

    # 1. 删除计划任务
    try {
        $taskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($taskExists) {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false | Out-Null
            Write-Host "已删除计划任务: $taskName" -ForegroundColor Green
        } else {
            Write-Host "计划任务不存在: $taskName" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "删除计划任务失败: $_" -ForegroundColor Red
    }

    # 2. 删除安装目录
    if (Test-Path $installPath) {
        try {
            Remove-Item -Path $installPath -Recurse -Force
            Write-Host "已删除安装目录: $installPath" -ForegroundColor Green
        } catch {
            Write-Host "删除安装目录失败: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "安装目录不存在: $installPath" -ForegroundColor Yellow
    }

    # 如果是完整卸载才恢复语言设置
    if (-not $Partial) {
        # 3. 自动恢复语言设置
        try {
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "ACP" -Value "936"
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "OEMCP" -Value "936"
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "MACCP" -Value "936"
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "ANSI" -Value "936"
            Write-Host "已恢复默认设置(All 936 zh_CN)" -ForegroundColor Green
        } catch {
            Write-Host "恢复设置失败: $_" -ForegroundColor Red
        }
    }

    if (-not $Partial) {
        # 退出脚本
        Write-Host "感谢您的使用,下次再见" -ForegroundColor Cyan
        Start-Sleep -Seconds 3
        Invoke-RestartPrompt -operation "卸载"
    }
}

# 安装功能封装为函数
function Install-DailyHeart {
    try {
        $scriptDir = $PSScriptRoot
        $sourceFolder = "$scriptDir\DailyHeart"
        $destinationPath = "C:\Program Files (x86)\DailyHeart"

        if (-not (Test-Path $sourceFolder -PathType Container)) {
            Write-Host "请检查完整性" -ForegroundColor Red
            Start-Sleep -Seconds 5
            Exit 1
        }

        Install-Module -Name BurntToast -Force

        # 若目标已存在，先删除
        if (Test-Path $destinationPath) {
            Remove-Item -Path $destinationPath -Recurse -Force
        }
        # 移动文件夹
        Move-Item -Path $sourceFolder -Destination $destinationPath -Force
        Write-Host "已成功移动" -ForegroundColor Green

        # 计划任务已存在时先删除
        if (Get-ScheduledTask -TaskName "DailyHeart" -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName "DailyHeart" -Confirm:$false | Out-Null
        }

        # 添加计划任务
        $scriptPath = "$destinationPath\main.ps1"
        $taskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        $taskTrigger = New-ScheduledTaskTrigger -AtLogOn -RandomDelay "00:00:03"
        $taskPrincipal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        try {
            $settings = New-ScheduledTaskSettingsSet -StartOnlyIfIdle:$false -DontStartIfOnBatteries:$false -WakeToRun
        } catch {
            Write-Host "部分计划任务参数不被当前系统支持，已自动降级" -ForegroundColor Yellow
            $settings = New-ScheduledTaskSettingsSet -WakeToRun
        }
        Register-ScheduledTask -TaskName "DailyHeart" -Action $taskAction -Trigger $taskTrigger -Settings $settings -Principal $taskPrincipal -Force | Out-Null
        Write-Host "已创建计划任务" -ForegroundColor Green
        
        # 设置控制台颜色
        $host.ui.RawUI.BackgroundColor = "Black"
        $host.ui.RawUI.ForegroundColor = "Red"
        Write-Host "本工具需要启用全局 UTF-8 支持才能正常运行(已知部分设备可以不启用)" -ForegroundColor Yellow
        Write-Host "警告：启用全局 UTF-8 支持可能会导致某些程序出现异常！" -ForegroundColor Red
        Write-Host "如果不确定后果或无法恢复系统，请不要继续！" -ForegroundColor Red

        do {
            $confirmInput = Read-Host "Y(es)/N(o)/S(kip)"
            switch ($confirmInput.ToUpper()) {
                "Y" {
                    try {
                        $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage"
                        Set-ItemProperty -Path $registryPath -Name "ACP" -Value "65001"
                        Set-ItemProperty -Path $registryPath -Name "OEMCP" -Value "65001"
                        Set-ItemProperty -Path $registryPath -Name "MACCP" -Value "65001"
                        Set-ItemProperty -Path $registryPath -Name "ANSI" -Value "65001"
                        Write-Host "已成功启用全局UTF-8支持！" -ForegroundColor Green
                        
                        # 安装完成
                        Write-Host "感谢使用DailyHeart,祝您生活愉快" -ForegroundColor Cyan
                        Start-Sleep -Seconds 5
                        Invoke-RestartPrompt -operation "安装"
                    } catch {
                        Write-Host "修改注册表时出错：$_" -ForegroundColor Red
                        # 回滚安装
                        Write-Host "正在回滚安装..." -ForegroundColor Yellow
                        Uninstall-DailyHeart -Partial $true
                        Start-Sleep -Seconds 5
                        Exit 1
                    }
                    break
                }
                "N" {
                    Write-Host "操作已取消，正在回滚安装..." -ForegroundColor Yellow
                    # 回滚安装
                    Uninstall-DailyHeart -Partial $true
                    Start-Sleep -Seconds 5
                    Exit 0
                }
                "S" {
                    # 安装完成
                    Write-Host "感谢使用DailyHeart,祝您生活愉快" -ForegroundColor Cyan
                    Start-Sleep -Seconds 5
                    Invoke-RestartPrompt -operation "安装"
                    break
                }
                default {
                    Write-Host "无效，请重新输入" -ForegroundColor Red
                }
            }
        }while ($true)
    } catch {
        Write-Host "安装过程中出错: $_" -ForegroundColor Red
        # 尝试执行部分卸载
        try {
            Write-Host "正在尝试回滚安装..." -ForegroundColor Yellow
            Uninstall-DailyHeart -Partial $true
        } catch {
            Write-Host "回滚安装失败: $_" -ForegroundColor Red
        }
        Start-Sleep -Seconds 5
        Exit 1
    }
}

# 恢复语言功能封装为函数
function Restore-LanguageSettings {
    Write-Host "正在恢复UTF-8支持设置(All 936 zh_CN)..." -ForegroundColor Yellow
    try {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "ACP" -Value "936"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "OEMCP" -Value "936"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "MACCP" -Value "936"
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "ANSI" -Value "936"
        Write-Host "已恢复默认设置" -ForegroundColor Green
        Invoke-RestartPrompt -operation "语言恢复"
    } catch {
        Write-Host "恢复设置失败: $_" -ForegroundColor Red
        Start-Sleep -Seconds 5
        Exit 1
    }
}

# 主菜单
do {
    $userInput = Read-Host "确定要安装 DailyHeart 吗?`nI(nstall)/U(ninstall)/R(estore Language)/E(xit)"
    switch -Regex ($userInput.ToUpper()) {
        "^I$" { Install-DailyHeart; break }
        "^U$" { Uninstall-DailyHeart; break }
        "^R$" { Restore-LanguageSettings; break }
        "^E$" {
            Write-Host "操作已取消。" -ForegroundColor Yellow
            Start-Sleep -Seconds 5
            Exit 0
        }
        default {
            Write-Host "无效，请重新输入" -ForegroundColor Red
        }
    }
} while ($true)