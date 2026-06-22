param(
    [string]$Python = "python",
    [switch]$Console
)

$ErrorActionPreference = "Stop"

$Root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$DistRoot = Join-Path $Root "dist"
$BuildRoot = Join-Path $Root "build"
New-Item -ItemType Directory -Path $DistRoot, $BuildRoot -Force | Out-Null

& $Python -m PyInstaller --version | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller is not installed for '$Python'. Run: $Python -m pip install -e .[build]"
}

$BuildInfoPath = Join-Path $Root "src\md2doc\build_info.py"
Write-Host "Generating build_info.py at: $BuildInfoPath"
$BuildTime = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')
$BuildInfoContent = "BUILD_TIME = `"$BuildTime`"`n"
[System.IO.File]::WriteAllText($BuildInfoPath, $BuildInfoContent)

$windowMode = if ($Console) { "--console" } else { "--windowed" }

& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --name md2doc `
    --specpath $BuildRoot `
    $windowMode `
    --paths (Join-Path $Root "src") `
    --collect-all markitdown `
    --collect-all playwright `
    (Join-Path $Root "scripts\pyinstaller_entry.py")
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed."
}

Write-Host "Created exe: $(Join-Path $DistRoot 'md2doc.exe')"
