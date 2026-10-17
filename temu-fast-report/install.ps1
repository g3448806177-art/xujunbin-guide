param(
    # 报单名：写入采购单 F 列并作为文件名前缀。不传时安装过程中询问。
    [string]$Reporter
)

$ErrorActionPreference = 'Stop'

$skillName = 'temu-fast-report'
$sourceSkill = Join-Path $PSScriptRoot $skillName
if (-not (Test-Path -LiteralPath (Join-Path $sourceSkill 'SKILL.md') -PathType Leaf)) {
    throw 'Skill files are missing. Extract the whole ZIP before running this installer.'
}

$desktop = [Environment]::GetFolderPath('DesktopDirectory')
if ([string]::IsNullOrWhiteSpace($desktop)) {
    throw 'Windows desktop path could not be resolved.'
}

$skillsRoot = Join-Path $env:USERPROFILE '.codex\skills'
$targetSkill = Join-Path $skillsRoot $skillName
$settingsFile = Join-Path $targetSkill 'config\settings.json'
$memory = Join-Path $desktop '报单记忆.xlsx'
$attributeDir = Join-Path $desktop '店铺属性库'
$attributeBook = Join-Path $attributeDir '店铺属性库.xlsx'

# Reuse the reporter name from a previous installation as the default answer.
$previousReporter = ''
if (Test-Path -LiteralPath $settingsFile -PathType Leaf) {
    try {
        $previous = Get-Content -LiteralPath $settingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($previous.reporter) { $previousReporter = ([string]$previous.reporter).Trim() }
    } catch {
        $previousReporter = ''
    }
}

$invalidChars = [IO.Path]::GetInvalidFileNameChars()
function Test-ReporterName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name.Length -gt 20) { return $false }
    return ($Name.IndexOfAny($invalidChars) -lt 0)
}

$Reporter = "$Reporter".Trim()
while (-not (Test-ReporterName $Reporter)) {
    if ($Reporter) {
        Write-Host '报单名最多20个字，且不能包含 \ / : * ? " < > | 等字符，请重新输入。' -ForegroundColor Yellow
    }
    if ($previousReporter) {
        $answer = Read-Host "请输入报单名（直接回车沿用：${previousReporter}）"
        if ([string]::IsNullOrWhiteSpace($answer)) { $answer = $previousReporter }
    } else {
        $answer = Read-Host '请输入报单名（会写入采购单F列和文件名，例如：张三）'
    }
    $Reporter = "$answer".Trim()
}

New-Item -ItemType Directory -Path $skillsRoot -Force | Out-Null
New-Item -ItemType Directory -Path $attributeDir -Force | Out-Null

# Never replace manually maintained workbooks.
if (-not (Test-Path -LiteralPath $memory)) {
    Copy-Item -LiteralPath (Join-Path $sourceSkill 'assets\空白报单记忆.xlsx') -Destination $memory
}
if (-not (Test-Path -LiteralPath $attributeBook)) {
    Copy-Item -LiteralPath (Join-Path $sourceSkill 'assets\空白店铺属性库.xlsx') -Destination $attributeBook
}

$backup = $null
if (Test-Path -LiteralPath $targetSkill) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backup = "$targetSkill.backup-$stamp"
    $counter = 1
    while (Test-Path -LiteralPath $backup) {
        $backup = "$targetSkill.backup-$stamp-$counter"
        $counter++
    }
    Move-Item -LiteralPath $targetSkill -Destination $backup
}

try {
    New-Item -ItemType Directory -Path $targetSkill -Force | Out-Null
    Copy-Item -Path (Join-Path $sourceSkill '*') -Destination $targetSkill -Recurse -Force
    New-Item -ItemType Directory -Path (Split-Path -Parent $settingsFile) -Force | Out-Null
    $config = [ordered]@{
        reporter = $Reporter
        desktop = $desktop
        memory = $memory
        attribute_dir = $attributeDir
    } | ConvertTo-Json -Compress
    [IO.File]::WriteAllText($settingsFile, $config, [Text.UTF8Encoding]::new($false))
} catch {
    if (Test-Path -LiteralPath $targetSkill) {
        Remove-Item -LiteralPath $targetSkill -Recurse -Force
    }
    if ($backup -and (Test-Path -LiteralPath $backup)) {
        Move-Item -LiteralPath $backup -Destination $targetSkill
    }
    throw
}

Write-Host ''
Write-Host "已安装：${targetSkill}"
Write-Host "报单名：${Reporter}"
Write-Host "报单记忆表：${memory}"
Write-Host "店铺属性库：${attributeDir}"
if ($backup) { Write-Host "旧版本已备份到：${backup}" }
Write-Host '桌面上已有的记忆表和属性库不会被覆盖。'
Write-Host '重启 Codex 后输入 $temu-fast-report 使用。'
