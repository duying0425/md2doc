param(
    [Parameter(Mandatory=$true)]
    [string]$Version
)

$ErrorActionPreference = "Stop"

$Root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$Python = Join-Path $Root ".venv\Scripts\python.exe"

# 1. Validate that python exists (use virtual environment if present, otherwise fallback to system python)
if (-not (Test-Path $Python)) {
    $Python = "python"
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

# 4. Git Commit version changes
Write-Host "Committing changes to Git..."
git add $InitPath $PyprojectPath
git commit -m "bump: version $Version"

# 5. Tag the commit
Write-Host "Tagging release as v$Version..."
git tag "v$Version"

# 6. Push branch and tag
Write-Host "Pushing to remote..."
git push origin main
git push origin "v$Version"

Write-Host "Version v$Version pushed successfully! GitHub Actions will now automatically build the executable and publish the Release." -ForegroundColor Green

