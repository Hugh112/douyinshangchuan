[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProjectRoot = [System.IO.Path]::GetFullPath($PSScriptRoot)
$Version = '3.0.3'
$MainDist = Join-Path $ProjectRoot 'dist\DouyinPublisher'
$UpdateZip = Join-Path $ProjectRoot "packages\DouyinPublisher_Update_v$Version.zip"
$InnoCompiler = 'C:\Users\1\AppData\Local\Programs\Inno Setup 6\ISCC.exe'
$IconPath = Join-Path $ProjectRoot 'app\assets\app_logo.ico'
$AssetsPath = Join-Path $ProjectRoot 'app\assets'
$AppPath = Join-Path $ProjectRoot 'app'
$VersionInfoPath = Join-Path $ProjectRoot 'build_tools\version_info.txt'

function Remove-BuildTarget([string]$TargetPath) {
    $full = [System.IO.Path]::GetFullPath($TargetPath)
    $prefix = $ProjectRoot.TrimEnd('\') + '\'
    if (-not $full.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean a path outside the workspace: $full"
    }
    if (Test-Path -LiteralPath $full) {
        Remove-Item -LiteralPath $full -Recurse -Force
    }
}

Set-Location -LiteralPath $ProjectRoot
Remove-BuildTarget (Join-Path $ProjectRoot 'build')
Remove-BuildTarget (Join-Path $ProjectRoot 'dist')
Remove-BuildTarget (Join-Path $ProjectRoot 'release')
Remove-BuildTarget (Join-Path $ProjectRoot 'DouyinPublisher.spec')
if (Test-Path -LiteralPath $UpdateZip) {
    Remove-Item -LiteralPath $UpdateZip -Force
}

py -m PyInstaller --noconfirm --clean --windowed --onedir `
    --name DouyinPublisher `
    --icon $IconPath `
    --version-file $VersionInfoPath `
    --specpath build\main-spec `
    --paths $AppPath `
    --add-data "$AssetsPath;assets" `
    --collect-all playwright `
    --hidden-import pandas `
    --hidden-import openpyxl `
    --hidden-import xlrd `
    app\app_main.py
if ($LASTEXITCODE -ne 0) { throw 'Main executable build failed.' }

py -m PyInstaller --noconfirm --clean --windowed --onefile `
    --name DouyinPublisherUpdater `
    --icon $IconPath `
    --distpath $MainDist `
    --workpath build\updater `
    --specpath build\updater-spec `
    app\online_updater.py
if ($LASTEXITCODE -ne 0) { throw 'Updater build failed.' }

$staging = Join-Path $ProjectRoot 'build\update-staging'
Remove-BuildTarget $staging
New-Item -ItemType Directory -Path $staging | Out-Null
Copy-Item -Path (Join-Path $MainDist '*') -Destination $staging -Recurse -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot 'release_transition\Start_Douyin_Publisher.vbs') -Destination $staging -Force
Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $UpdateZip -CompressionLevel Optimal

if (-not (Test-Path -LiteralPath $InnoCompiler)) {
    throw "Inno Setup compiler was not found: $InnoCompiler"
}
& $InnoCompiler (Join-Path $ProjectRoot 'installer\DouyinPublisher.iss')
if ($LASTEXITCODE -ne 0) { throw 'Installer build failed.' }

$zipHash = (Get-FileHash -LiteralPath $UpdateZip -Algorithm SHA256).Hash.ToLowerInvariant()
$zipSize = (Get-Item -LiteralPath $UpdateZip).Length
$installer = Join-Path $ProjectRoot "release\DouyinPublisher_Setup_v$Version.exe"
$installerHash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
Remove-BuildTarget (Join-Path $ProjectRoot 'build')
Remove-BuildTarget (Join-Path $ProjectRoot 'app\__pycache__')
Write-Host "UPDATE_ZIP=$UpdateZip"
Write-Host "UPDATE_SHA256=$zipHash"
Write-Host "UPDATE_SIZE=$zipSize"
Write-Host "INSTALLER=$installer"
Write-Host "INSTALLER_SHA256=$installerHash"
