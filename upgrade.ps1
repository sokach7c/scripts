# 扫描 Windows 电脑中的论文/毕业设计类文档（不读取文件内容）
# 依据：目录、文件名、大小、类型、附加标识 进行评分并过滤

[CmdletBinding()]
param(
    # 额外扫描根路径（默认会扫描各用户桌面、文档、下载及常见盘符根下的一级目录）
    [string[]]$ScanRoots = @(),

    # 最低得分阈值，低于此分数的结果不输出
    [int]$MinScore = 35,

    # 单文件大小下限（字节），低于此值直接跳过，不参与评分
    [long]$MinFileSizeBytes = 100KB,

    # 单文件大小上限（字节），过大多为安装包/镜像
    [long]$MaxFileSizeBytes = 80MB,

    # 是否递归子目录（建议保持 true）
    [switch]$Recurse = $true
)

$ErrorActionPreference = 'SilentlyContinue'

#region 配置：关键词与排除规则

# 路径关键词（命中加分）
$PathKeywords = @(
    '毕业', '论文', '设计', '开题', '答辩', '文献', '学位',
    'thesis', 'dissertation', 'graduate', 'graduation', 'paper', 'design'
)

# 文件名关键词（命中加分，权重更高）
$NameKeywords = @(
    @{ Pattern = '毕业论文|学位论文|本科论文|硕士论文|博士论文|博硕论文|研究生论文'; Score = 25 },
    @{ Pattern = '毕业设计|课程设计|毕业实习|实践报告|设计说明书'; Score = 22 },
    @{ Pattern = '开题报告|文献综述|任务书|中期检查|中期报告|答辩|PPT大纲'; Score = 20 },
    @{ Pattern = '终稿|定稿|终版|正式版|完整版|答辩版|提交版|送审版|存档版|清稿|核稿'; Score = 18 },
    @{ Pattern = '论文|设计说明|设计报告|设计论文|research|thesis|dissertation|paper'; Score = 15 },
    @{ Pattern = '选题|绪论|引言|概述|综述|摘要|abstract|acknowledgement|致谢|参考文献|文献翻译'; Score = 10 },
    @{ Pattern = '学院|专业|\d{2,4}届|\d{2,4}级|学号|指导教师|导师'; Score = 6 },
    @{ Pattern = 'v\d+|final|submission|draft\d|修订|修改版|改\d|第\d版'; Score = 8 }
)

# 仅扫描 Word 文档
$DocExtensions = @{
    '.docx' = 20
    '.doc'  = 18
}

# 文件名中的负面模式（减分）
$NegativeNamePatterns = @(
    @{ Pattern = '安装|setup|installer|uninstall|驱动|driver|readme'; Score = -30 },
    @{ Pattern = '模板\.(doc|docx)$|sample|example|示例|空白|blank'; Score = -15 },
    @{ Pattern = '^\~\$|\.tmp$|\.bak$|\.old$|副本|copy|复件'; Score = -40 }
)

# 绝对跳过的目录名（不进入）
$SkipDirNames = @(
    'node_modules', 'npm-cache', '.npm', '.yarn', 'yarn-cache', 'pnpm-store',
    '.git', '.svn', '.hg', '.idea', '.vscode', '.vs', 'vendor', 'packages',
    '__pycache__', '.pytest_cache', '.mypy_cache', '.venv', 'venv', 'env',
    'target', 'build', 'dist', 'out', '.gradle', '.nuget', 'bower_components',
    'Windows', 'WinSxS', 'System32', 'SysWOW64', 'WinRE', 'Recovery',
    'Program Files', 'Program Files (x86)', 'ProgramData',
    '$Recycle.Bin', 'System Volume Information', 'AppData',
    'Microsoft', 'WindowsApps', 'WinSAT', 'Prefetch', 'Installer',
    'servicing', 'assembly', 'Boot', 'Logs', 'Temp', 'tmp',
    'Intel', 'AMD', 'NVIDIA', 'Drivers', 'DriverStore',
    'Cache', 'Caches', 'CachedData', 'Code Cache', 'GPUCache',
    'Steam', 'steamapps', 'Epic Games', 'Origin', 'Battle.net'
)

# 路径片段负面关键词（整条路径命中则跳过文件）
$SkipPathFragments = @(
    '\Windows\', '\WinSxS\', '\System32\', '\SysWOW64\',
    '\Program Files\', '\Program Files (x86)\', '\ProgramData\Microsoft\',
    '\node_modules\', '\npm-cache\', '\.git\', '\AppData\Local\Temp\',
    '\AppData\Local\Packages\', '\AppData\Local\Microsoft\Windows\',
    '\$Recycle.Bin\', '\System Volume Information\'
)

#endregion

#region 辅助函数

function Get-DefaultScanRoots {
    $roots = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($profile in @($env:USERPROFILE, $env:PUBLIC)) {
        if (-not $profile) { continue }

        foreach ($sub in @('Desktop', '桌面', 'Documents', '文档', 'Downloads', '下载')) {
            $p = Join-Path $profile $sub
            if (Test-Path $p) { [void]$roots.Add($p) }
        }

        Get-ChildItem -Path $profile -Directory -Filter 'OneDrive*' -ErrorAction SilentlyContinue |
            ForEach-Object { [void]$roots.Add($_.FullName) }
    }

    foreach ($drive in Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue) {
        $root = $drive.Root
        if ($root -match '^[A-Z]:\\$') {
            try {
                Get-ChildItem -Path $root -Directory -ErrorAction SilentlyContinue |
                    Where-Object {
                        $n = $_.Name
                        -not ($SkipDirNames | Where-Object { $n -ieq $_ })
                    } |
                    ForEach-Object { [void]$roots.Add($_.FullName) }
            } catch { }
        }
    }

    return @($roots)
}

function Test-ShouldSkipDirectory {
    param([string]$DirPath)

    $name = Split-Path $DirPath -Leaf
    foreach ($skip in $SkipDirNames) {
        if ($name -ieq $skip) { return $true }
    }
    return $false
}

function Test-ShouldSkipFilePath {
    param([string]$FilePath)

    foreach ($frag in $SkipPathFragments) {
        if ($FilePath -like "*$frag*") { return $true }
    }
    return $false
}

function Get-DesktopBonus {
    param([string]$FilePath)

    $normalized = $FilePath -replace '\\', '/'
    if ($normalized -match '/Desktop/|/桌面/') { return 15 }
    if ($normalized -match '/Documents/|/文档/|/Downloads/|/下载/') { return 8 }
    return 0
}

function Get-PathKeywordScore {
    param([string]$FilePath)

    $score = 0
    $lower = $FilePath.ToLowerInvariant()
    foreach ($kw in $PathKeywords) {
        if ($lower.Contains($kw.ToLowerInvariant())) {
            $score += 10
        }
    }
    return [Math]::Min($score, 30)
}

function Get-NameKeywordScore {
    param([string]$FileName)

    $score = 0
    foreach ($rule in $NameKeywords) {
        if ($FileName -match $rule.Pattern) {
            $score += $rule.Score
        }
    }
    foreach ($rule in $NegativeNamePatterns) {
        if ($FileName -match $rule.Pattern) {
            $score += $rule.Score
        }
    }
    return $score
}

function Get-ExtensionScore {
    param([string]$Extension)

    $ext = $Extension.ToLowerInvariant()
    if (-not $ext.StartsWith('.')) { $ext = ".$ext" }
    if ($DocExtensions.ContainsKey($ext)) {
        return $DocExtensions[$ext]
    }
    return 0
}

function Get-SizeScore {
    param([long]$Length)

    if ($Length -gt $MaxFileSizeBytes) { return -15 }

    # 论文常见体量：约 200KB ~ 25MB 加分
    if ($Length -ge 200KB -and $Length -le 25MB) { return 12 }
    if ($Length -ge $MinFileSizeBytes -and $Length -le 40MB) { return 6 }
    return 0
}

function Get-FormatBonus {
    param([string]$FilePath)

    $bonus = 0
    $dirName = Split-Path (Split-Path $FilePath -Parent) -Leaf

    # 目录名像论文文件夹
    if ($dirName -match '论文|毕业|thesis|dissertation|设计|开题|答辩|学位') {
        $bonus += 8
    }

    return $bonus
}

function Get-ThesisFileScore {
    param([System.IO.FileInfo]$File)

    $ext = $File.Extension
    if ($ext -notin $DocExtensions.Keys) { return $null }
    if ($File.LongLength -lt $MinFileSizeBytes) { return $null }

    $extScore = Get-ExtensionScore -Extension $ext
    $path = $File.FullName
    if (Test-ShouldSkipFilePath -FilePath $path) {
        return $null
    }

    $breakdown = [ordered]@{
        Extension   = $extScore
        DesktopDoc  = (Get-DesktopBonus -FilePath $path)
        PathKeyword = (Get-PathKeywordScore -FilePath $path)
        NameKeyword = (Get-NameKeywordScore -FileName $File.Name)
        Size        = (Get-SizeScore -Length $File.LongLength)
        FormatBonus = (Get-FormatBonus -FilePath $path)
    }

    $total = ($breakdown.Values | Measure-Object -Sum).Sum

    [PSCustomObject]@{
        Path       = $path
        Name       = $File.Name
        Extension  = $ext
        SizeBytes  = $File.LongLength
        SizeHuman  = if ($File.LongLength -ge 1MB) {
            '{0:N2} MB' -f ($File.LongLength / 1MB)
        } elseif ($File.LongLength -ge 1KB) {
            '{0:N1} KB' -f ($File.LongLength / 1KB)
        } else {
            '{0} B' -f $File.LongLength
        }
        Modified   = $File.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
        Score      = $total
        Breakdown  = ($breakdown.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '
    }
}

function Get-ScannableFiles {
    param(
        [string]$Root,
        [bool]$DoRecurse
    )

    $options = @{
        File  = $true
        ErrorAction = 'SilentlyContinue'
    }
    if ($DoRecurse) {
        $options['Recurse'] = $true
    }

    try {
        Get-ChildItem -Path $Root @options | Where-Object {
            -not $_.PSIsContainer -and
            ($_.Extension -in $DocExtensions.Keys) -and
            ($_.Length -ge $MinFileSizeBytes)
        }
    } catch {
        @()
    }
}

#endregion

#region 主流程

Write-Host ''
Write-Host '========================================' -ForegroundColor Cyan
Write-Host ' 论文/毕业设计 Word 文档扫描（.doc/.docx，不读内容）' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan
Write-Host "最低得分阈值: $MinScore"
Write-Host "文件类型: .doc / .docx"
Write-Host "大小范围: >= $MinFileSizeBytes（过小直接跳过）~ $MaxFileSizeBytes 字节"
Write-Host ''

$allRoots = if ($ScanRoots.Count -gt 0) { $ScanRoots } else { Get-DefaultScanRoots }
Write-Host "扫描根路径 ($($allRoots.Count) 个):" -ForegroundColor Yellow
$allRoots | ForEach-Object { Write-Host "  - $_" }
Write-Host ''

$results = [System.Collections.Generic.List[object]]::new()
$scannedFiles = 0
$skippedDirs = 0

foreach ($root in $allRoots) {
    if (-not (Test-Path $root)) { continue }

    Write-Host "正在扫描: $root ..." -ForegroundColor DarkGray

    if ($Recurse) {
        # 手动 BFS，便于按目录名提前剪枝
        $queue = [System.Collections.Queue]::new()
        $queue.Enqueue($root)

        while ($queue.Count -gt 0) {
            $current = $queue.Dequeue()
            if (Test-ShouldSkipDirectory -DirPath $current) {
                $skippedDirs++
                continue
            }

            try {
                foreach ($item in Get-ChildItem -Path $current -Force -ErrorAction SilentlyContinue) {
                    if ($item.PSIsContainer) {
                        if (-not (Test-ShouldSkipDirectory -DirPath $item.FullName)) {
                            $queue.Enqueue($item.FullName)
                        } else {
                            $skippedDirs++
                        }
                    } elseif ($item.Extension -in $DocExtensions.Keys -and $item.Length -ge $MinFileSizeBytes) {
                        $scannedFiles++
                        $scored = Get-ThesisFileScore -File $item
                        if ($null -ne $scored -and $scored.Score -ge $MinScore) {
                            $results.Add($scored)
                        }
                    }
                }
            } catch { }
        }
    } else {
        foreach ($file in Get-ScannableFiles -Root $root -DoRecurse:$false) {
            $scannedFiles++
            $scored = Get-ThesisFileScore -File $file
            if ($null -ne $scored -and $scored.Score -ge $MinScore) {
                $results.Add($scored)
            }
        }
    }
}

$sorted = $results | Sort-Object Score -Descending, SizeBytes -Descending

Write-Host ''
Write-Host '========================================' -ForegroundColor Green
Write-Host " 扫描完成" -ForegroundColor Green
Write-Host '========================================' -ForegroundColor Green
Write-Host "候选文档扩展名命中: $scannedFiles 个文件"
Write-Host "剪枝跳过目录: $skippedDirs 个"
Write-Host "得分 >= $MinScore 的结果: $($sorted.Count) 个"
Write-Host ''

if ($sorted.Count -eq 0) {
    Write-Host '未找到符合条件的论文类文档。可尝试：' -ForegroundColor Yellow
    Write-Host '  - 降低 -MinScore（例如 25）'
    Write-Host '  - 通过 -ScanRoots 指定额外路径'
    exit 0
}

$index = 1
foreach ($item in $sorted) {
    Write-Host ("[{0}] 得分 {1}" -f $index, $item.Score) -ForegroundColor White
    Write-Host ("     文件: {0}" -f $item.Name)
    Write-Host ("     路径: {0}" -f $item.Path)
    Write-Host ("     类型: {0}  |  大小: {1}  |  修改: {2}" -f $item.Extension, $item.SizeHuman, $item.Modified)
    Write-Host ("     评分明细: {0}" -f $item.Breakdown) -ForegroundColor DarkGray
    Write-Host ''
    $index++
}

# 汇总表（便于复制）
Write-Host '--- 汇总表 ---' -ForegroundColor Cyan
$sorted | Format-Table -Property Score, Name, Extension, SizeHuman, Modified, Path -AutoSize | Out-String | Write-Host

#endregion
