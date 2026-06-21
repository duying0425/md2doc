param(
    [Parameter(Mandatory=$true)]
    [string]$Version
)

$ErrorActionPreference = "Stop"

$Root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$Python = Join-Path $Root ".venv\Scripts\python.exe"

# 1. Validate that the virtual environment exists
if (-not (Test-Path $Python)) {
    throw "Virtual environment Python not found at: $Python"
}

# 2. Validate Git working tree is clean
$gitStatus = git status --porcelain
if ($gitStatus) {
    throw "Git working tree is not clean. Please commit or stash changes before releasing."
}

# 3. Bump version in files using Python to avoid encoding issues
$InitPath = Join-Path $Root "src\md2doc\__init__.py"
$PyprojectPath = Join-Path $Root "pyproject.toml"

Write-Host "Bumping version to $Version in files..."
& $Python -c "import re; p = r'$InitPath'; c = open(p, 'r', encoding='utf-8').read(); c = re.sub(r'__version__\s*=\s*[\x22\x27][^\x22\x27]+[\x22\x27]', f'__version__ = \x22$Version\x22', c); open(p, 'w', encoding='utf-8').write(c)"
& $Python -c "import re; p = r'$PyprojectPath'; c = open(p, 'r', encoding='utf-8').read(); c = re.sub(r'version\s*=\s*[\x22\x27][^\x22\x27]+[\x22\x27]', f'version = \x22$Version\x22', c); open(p, 'w', encoding='utf-8').write(c)"

# 4. Trigger build process (which writes build_info.py and compiles the executable)
Write-Host "Building executable..."
powershell -File (Join-Path $PSScriptRoot "build_exe.ps1") -Python $Python
if ($LASTEXITCODE -ne 0) {
    throw "Build failed."
}

# 5. Git Commit version changes and build_info.py
Write-Host "Committing changes to Git..."
$BuildInfoPath = Join-Path $Root "src\md2doc\build_info.py"
git add $InitPath $PyprojectPath $BuildInfoPath
git commit -m "bump: version $Version and build time"

# 6. Tag the commit
Write-Host "Tagging release as v$Version..."
git tag "v$Version"

# 7. Push branch and tag
Write-Host "Pushing to remote..."
git push origin main
git push origin "v$Version"

# 8. Create GitHub Release
Write-Host "Creating GitHub Release and uploading build artifact..."
$ExePath = Join-Path $Root "dist\md2doc.exe"
gh release create "v$Version" $ExePath --title "v$Version" --notes "Release v$Version"

Write-Host "Release v$Version published successfully!" -ForegroundColor Green
