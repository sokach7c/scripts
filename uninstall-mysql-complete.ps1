#requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch]$SkipUninstallExe,
    [switch]$DryRun
)

$ErrorActionPreference = 'SilentlyContinue'

$LogFile = Join-Path $env:TEMP ("mysql-uninstall-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".log")

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$time] [$Level] $Message"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

function Invoke-Step {
    param(
        [string]$Title,
        [scriptblock]$Action
    )
    Write-Log "========== $Title =========="
    try {
        & $Action
    }
    catch {
        Write-Log "$Title 执行失败：$($_.Exception.Message)" "ERROR"
    }
}

function Remove-RegistryKeyIfExists {
    param([string]$Path)

    if (Test-Path $Path) {
        if ($DryRun) {
            Write-Log "[预览] 将删除注册表项: $Path"
        } else {
            try {
                Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                Write-Log "已删除注册表项: $Path"
            }
            catch {
                Write-Log "删除注册表项失败: $Path，原因: $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

function Remove-FileSystemPathIfExists {
    param([string]$Path)

    if (Test-Path $Path) {
        if ($DryRun) {
            Write-Log "[预览] 将删除路径: $Path"
        } else {
            try {
                Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                Write-Log "已删除路径: $Path"
            }
            catch {
                Write-Log "删除路径失败: $Path，原因: $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

function Remove-MatchingChildren {
    param(
        [string]$ParentPath,
        [string]$NamePattern
    )

    if (Test-Path $ParentPath) {
        Get-ChildItem -Path $ParentPath -ErrorAction SilentlyContinue | Where-Object {
            $_.PSChildName -match $NamePattern
        } | ForEach-Object {
            Remove-RegistryKeyIfExists -Path $_.PSPath
        }
    }
}

function Stop-And-Delete-Service {
    param([string]$ServiceName)

    try {
        $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($null -ne $svc) {
            if ($svc.Status -ne 'Stopped') {
                if ($DryRun) {
                    Write-Log "[预览] 将停止服务: $ServiceName"
                } else {
                    Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
                    Write-Log "已停止服务: $ServiceName"
                }
            }

            if ($DryRun) {
                Write-Log "[预览] 将删除服务: $ServiceName"
            } else {
                sc.exe delete $ServiceName | Out-Null
                Write-Log "已删除服务: $ServiceName"
            }
        }
    }
    catch {
        Write-Log "处理服务失败: $ServiceName，原因: $($_.Exception.Message)" "ERROR"
    }
}

function Remove-FromMachinePath {
    param([string[]]$Keywords)

    $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    if ([string]::IsNullOrWhiteSpace($currentPath)) {
        return
    }

    $parts = $currentPath -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $newParts = @()

    foreach ($p in $parts) {
        $lower = $p.ToLowerInvariant()
        $matched = $false
        foreach ($keyword in $Keywords) {
            if ($lower -like "*$($keyword.ToLowerInvariant())*") {
                $matched = $true
                Write-Log "PATH 将移除项: $p"
                break
            }
        }
        if (-not $matched) {
            $newParts += $p
        }
    }

    $newPath = ($newParts -join ';')
    if ($DryRun) {
        Write-Log "[预览] 将更新系统 PATH"
    } else {
        [Environment]::SetEnvironmentVariable("Path", $newPath, "Machine")
        Write-Log "已更新系统 PATH"
    }
}

Write-Log "MySQL 完整彻底卸载开始"
Write-Log "日志文件: $LogFile"
Write-Log "DryRun 模式: $DryRun"
Write-Log "SkipUninstallExe: $SkipUninstallExe"

Invoke-Step "1. 停止并删除 MySQL 服务" {
    $candidateServiceNames = @(
        "MySQL",
        "MySQL80",
        "MySQL57",
        "MySQL56",
        "MySQL55",
        "mysql",
        "mysqld"
    )

    foreach ($name in $candidateServiceNames) {
        Stop-And-Delete-Service -ServiceName $name
    }

    Get-Service -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match '(?i)^mysql' -or $_.DisplayName -match '(?i)mysql'
    } | ForEach-Object {
        Stop-And-Delete-Service -ServiceName $_.Name
    }
}

Invoke-Step "2. 卸载注册表中的 MySQL / Oracle MySQL 程序" {
    if (-not $SkipUninstallExe) {
        $uninstallRoots = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )

        foreach ($root in $uninstallRoots) {
            if (Test-Path $root) {
                Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
                    $item = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                    $displayName = $item.DisplayName
                    $uninstallString = $item.UninstallString

                    if ($displayName -and ($displayName -match '(?i)mysql' -or $displayName -match '(?i)oracle.*mysql')) {
                        Write-Log "检测到卸载项: $displayName"

                        if (-not [string]::IsNullOrWhiteSpace($uninstallString)) {
                            if ($DryRun) {
                                Write-Log "[预览] 将执行卸载命令: $uninstallString"
                            } else {
                                try {
                                    if ($uninstallString -match '(?i)msiexec\.exe') {
                                        $args = $uninstallString -replace '(?i)^.*?msiexec\.exe\s*', ''
                                        Start-Process -FilePath "msiexec.exe" -ArgumentList "$args /qn" -Wait -NoNewWindow
                                    } else {
                                        Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $uninstallString -Wait -NoNewWindow
                                    }
                                    Write-Log "已尝试卸载: $displayName"
                                }
                                catch {
                                    Write-Log "卸载失败: $displayName，原因: $($_.Exception.Message)" "ERROR"
                                }
                            }
                        }
                    }
                }
            }
        }
    } else {
        Write-Log "已跳过卸载程序步骤"
    }
}

Invoke-Step "3. 删除常见安装目录、数据目录、用户目录" {
    $paths = @(
        "C:\Program Files\MySQL",
        "C:\Program Files (x86)\MySQL",
        "C:\ProgramData\MySQL",
        "C:\ProgramData\MySQL\MySQL Installer",
        "C:\MySQL",
        "$env:LOCALAPPDATA\MySQL",
        "$env:APPDATA\MySQL"
    )

    foreach ($p in $paths) {
        Remove-FileSystemPathIfExists -Path $p
    }

    # 清理 Package Cache 中名称带 mysql 的目录/文件
    $packageCache = "C:\ProgramData\Package Cache"
    if (Test-Path $packageCache) {
        Get-ChildItem $packageCache -Recurse -Force -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -match '(?i)mysql'
        } | Sort-Object FullName -Descending | ForEach-Object {
            if ($_.PSIsContainer) {
                Remove-FileSystemPathIfExists -Path $_.FullName
            } else {
                if ($DryRun) {
                    Write-Log "[预览] 将删除文件: $($_.FullName)"
                } else {
                    try {
                        Remove-Item -Path $_.FullName -Force -ErrorAction Stop
                        Write-Log "已删除文件: $($_.FullName)"
                    }
                    catch {
                        Write-Log "删除文件失败: $($_.FullName)，原因: $($_.Exception.Message)" "ERROR"
                    }
                }
            }
        }
    }
}

Invoke-Step "4. 清理系统环境变量" {
    Remove-FromMachinePath -Keywords @("mysql")
    if ($DryRun) {
        Write-Log "[预览] 将删除系统环境变量 MYSQL_HOME"
    } else {
        [Environment]::SetEnvironmentVariable("MYSQL_HOME", $null, "Machine")
        Write-Log "已删除系统环境变量 MYSQL_HOME"
    }
}

Invoke-Step "5. 清理常见软件注册表残留" {
    $regPaths = @(
        "HKLM:\SOFTWARE\MySQL",
        "HKLM:\SOFTWARE\MySQL AB",
        "HKCU:\SOFTWARE\MySQL",
        "HKLM:\SOFTWARE\WOW6432Node\MySQL",
        "HKLM:\SOFTWARE\WOW6432Node\MySQL AB",
        "HKLM:\SOFTWARE\Oracle\MySQL"
    )

    foreach ($rp in $regPaths) {
        Remove-RegistryKeyIfExists -Path $rp
    }

    # Oracle 根节点下仅删除和 MySQL 有关的子项
    $oracleRoot = "HKLM:\SOFTWARE\Oracle"
    if (Test-Path $oracleRoot) {
        Remove-MatchingChildren -ParentPath $oracleRoot -NamePattern '(?i)mysql'
    }
}

Invoke-Step "6. 清理 SYSTEM 下 ControlSet / CurrentControlSet 中的 MySQL 服务注册表项" {
    $controlSets = Get-ChildItem "HKLM:\SYSTEM" -ErrorAction SilentlyContinue | Where-Object {
        $_.PSChildName -match '^ControlSet\d{3}$|^CurrentControlSet$'
    } | Select-Object -ExpandProperty PSChildName

    foreach ($cs in $controlSets) {
        Write-Log "处理注册表集: $cs"

        $servicesPath = "HKLM:\SYSTEM\$cs\Services"
        if (Test-Path $servicesPath) {
            Get-ChildItem $servicesPath -ErrorAction SilentlyContinue | Where-Object {
                $_.PSChildName -match '(?i)^mysql'
            } | ForEach-Object {
                Remove-RegistryKeyIfExists -Path $_.PSPath
            }
        }
    }
}

Invoke-Step "7. 清理 EventLog 中的 MySQL 日志源注册表项" {
    $controlSets = Get-ChildItem "HKLM:\SYSTEM" -ErrorAction SilentlyContinue | Where-Object {
        $_.PSChildName -match '^ControlSet\d{3}$|^CurrentControlSet$'
    } | Select-Object -ExpandProperty PSChildName

    foreach ($cs in $controlSets) {
        $eventLogPath = "HKLM:\SYSTEM\$cs\Services\EventLog\Application"
        if (Test-Path $eventLogPath) {
            Get-ChildItem $eventLogPath -ErrorAction SilentlyContinue | Where-Object {
                $_.PSChildName -match '(?i)mysql'
            } | ForEach-Object {
                Remove-RegistryKeyIfExists -Path $_.PSPath
            }
        }
    }
}

Invoke-Step "8. 清理防火墙规则" {
    $rules = Get-NetFirewallRule -ErrorAction SilentlyContinue | Where-Object {
        $_.DisplayName -match '(?i)mysql'
    }

    foreach ($rule in $rules) {
        if ($DryRun) {
            Write-Log "[预览] 将删除防火墙规则: $($rule.DisplayName)"
        } else {
            try {
                Remove-NetFirewallRule -Name $rule.Name -ErrorAction Stop
                Write-Log "已删除防火墙规则: $($rule.DisplayName)"
            }
            catch {
                Write-Log "删除防火墙规则失败: $($rule.DisplayName)，原因: $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

Invoke-Step "9. 检查 3306 端口占用" {
    $connections = Get-NetTCPConnection -LocalPort 3306 -ErrorAction SilentlyContinue
    if ($connections) {
        foreach ($conn in $connections) {
            Write-Log "端口 3306 仍被占用: OwningProcess=$($conn.OwningProcess), State=$($conn.State)" "WARN"
        }
    } else {
        Write-Log "端口 3306 未发现占用"
    }
}

Invoke-Step "10. 最终检查" {
    $leftServices = Get-Service -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match '(?i)^mysql' -or $_.DisplayName -match '(?i)mysql'
    }

    if ($leftServices) {
        foreach ($svc in $leftServices) {
            Write-Log "残留服务: $($svc.Name) / $($svc.DisplayName)" "WARN"
        }
    } else {
        Write-Log "未发现 MySQL 服务残留"
    }

    $checkPaths = @(
        "C:\Program Files\MySQL",
        "C:\Program Files (x86)\MySQL",
        "C:\ProgramData\MySQL",
        "C:\MySQL",
        "$env:LOCALAPPDATA\MySQL",
        "$env:APPDATA\MySQL"
    )

    foreach ($cp in $checkPaths) {
        if (Test-Path $cp) {
            Write-Log "残留目录: $cp" "WARN"
        }
    }
}

Write-Log "MySQL 完整彻底卸载结束"
Write-Log "日志文件位置: $LogFile"
Write-Host ""
Write-Host "执行完成。日志文件: $LogFile" -ForegroundColor Green
