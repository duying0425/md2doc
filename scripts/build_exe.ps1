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
    (Join-Path $Root "scripts\pyinstaller_entry.py")
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed."
}

Write-Host "Created exe: $(Join-Path $DistRoot 'md2doc.exe')"
