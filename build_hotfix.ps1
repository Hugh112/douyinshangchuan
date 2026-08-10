[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProjectRoot = [System.IO.Path]::GetFullPath($PSScriptRoot)
$Version = '3.0.4'
$BuildRoot = Join-Path $ProjectRoot 'build\hotfix'
$PayloadRoot = Join-Path $BuildRoot 'dist'
$PayloadDir = Join-Path $PayloadRoot 'DouyinPublisher'
$InnoCompiler = 'C:\Users\1\AppData\Local\Programs\Inno Setup 6\ISCC.exe'
$IconPath = Join-Path $ProjectRoot 'app\assets\app_logo.ico'
$AssetsPath = Join-Path $ProjectRoot 'app\assets'
$AppPath = Join-Path $ProjectRoot 'app'
$VersionInfoPath = Join-Path $ProjectRoot 'build_tools\version_info.txt'

function Remove-HotfixBuildTarget([string]$TargetPath) {
    $full = [System.IO.Path]::GetFullPath($TargetPath)
    $allowed = [System.IO.Path]::GetFullPath($BuildRoot)
    if (-not $full.StartsWith($allowed, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean outside the hotfix build directory: $full"
    }
    if (Test-Path -LiteralPath $full) {
        Remove-Item -LiteralPath $full -Recurse -Force
    }
}

Set-Location -LiteralPath $ProjectRoot
if (Test-Path -LiteralPath $BuildRoot) {
    $resolvedBuild = [System.IO.Path]::GetFullPath($BuildRoot)
    if ($resolvedBuild -ne [System.IO.Path]::GetFullPath((Join-Path $ProjectRoot 'build\hotfix'))) {
        throw "Unexpected hotfix build path: $resolvedBuild"
    }
    Remove-Item -LiteralPath $resolvedBuild -Recurse -Force
}

py -m PyInstaller --noconfirm --clean --windowed --onedir `
    --name DouyinPublisher `
    --icon $IconPath `
    --version-file $VersionInfoPath `
    --distpath $PayloadRoot `
    --workpath (Join-Path $BuildRoot 'work') `
    --specpath (Join-Path $BuildRoot 'spec') `
    --paths $AppPath `
    --add-data "$AssetsPath;assets" `
    --collect-all playwright `
    --hidden-import pandas `
    --hidden-import openpyxl `
    --hidden-import xlrd `
    app\app_main.py
if ($LASTEXITCODE -ne 0) { throw 'Hotfix main executable build failed.' }

$payloadExe = Join-Path $PayloadDir 'DouyinPublisher.exe'
if (-not (Test-Path -LiteralPath $payloadExe)) {
    throw "Hotfix payload was not generated: $payloadExe"
}
$fileVersion = (Get-Item -LiteralPath $payloadExe).VersionInfo.FileVersion
if ($fileVersion -ne "$Version.0") {
    throw "Hotfix payload version mismatch: $fileVersion"
}
if (-not (Test-Path -LiteralPath $InnoCompiler)) {
    throw "Inno Setup compiler was not found: $InnoCompiler"
}

& $InnoCompiler (Join-Path $ProjectRoot 'installer\DouyinPublisherHotfix.iss')
if ($LASTEXITCODE -ne 0) { throw 'Hotfix installer build failed.' }

$hotfix = Join-Path $ProjectRoot "release\DouyinPublisher_Hotfix_v$Version.exe"
$hotfixHash = (Get-FileHash -LiteralPath $hotfix -Algorithm SHA256).Hash.ToLowerInvariant()
$hotfixSize = (Get-Item -LiteralPath $hotfix).Length
Write-Host "HOTFIX=$hotfix"
Write-Host "HOTFIX_SHA256=$hotfixHash"
Write-Host "HOTFIX_SIZE=$hotfixSize"
