<#
.SYNOPSIS
    Bumps the version, commits version changes, tags, and pushes to remote Git repository.

.DESCRIPTION
    This script automates the git release process. It will:
    1. Detect the Python interpreter in the virtual environment.
    2. Ensure the working tree is clean.
    3. Bump the version in `pyproject.toml` and `src/md2doc/__init__.py`.
       If no version is specified, it automatically increments the patch version (e.g., x.y.z -> x.y.z+1).
    4. Commit the changes and tag it as `v<Version>`.
    5. Push the commit and the tag to origin.
       This pushes the tag `v<Version>` which triggers GitHub Actions to build the executable and create a GitHub Release.

.PARAMETER Version
    The target version number (e.g., 0.4.2). If omitted, the patch (third) component of the
    current version in `pyproject.toml` will be automatically incremented.

.EXAMPLE
    powershell -File scripts/release.ps1
    (Bumps patch version automatically, e.g. 0.4.1 -> 0.4.2, commits, tags, and pushes to remote)

.EXAMPLE
    powershell -File scripts/release.ps1 -Version 0.5.0
    (Updates version to 0.5.0, commits, tags, and pushes to remote)
#>
param(
    [Parameter(Mandatory=$false)]
    [string]$Version
)

$ErrorActionPreference = "Stop"

$Root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$Python = Join-Path $Root ".venv\Scripts\python.exe"

# 1. Validate that python exists (use virtual environment if present, otherwise fallback to system python)
if (-not (Test-Path $Python)) {
    $Python = "python"
}

# Auto-increment patch version if not specified
if (-not $Version) {
    $PyprojectPath = Join-Path $Root "pyproject.toml"
    $PyprojectContent = Get-Content -Path $PyprojectPath -Raw
    if ($PyprojectContent -match 'version\s*=\s*"([^"]+)"') {
        $CurrentVersion = $Matches[1]
        $parts = $CurrentVersion.Split('.')
        if ($parts.Length -eq 3) {
            $patch = [int]$parts[2] + 1
            $Version = "$($parts[0]).$($parts[1]).$patch"
            Write-Host "No version specified. Automatically incrementing patch version: $CurrentVersion -> $Version" -ForegroundColor Cyan
        } else {
            throw "Could not parse 3-part semver version from '$CurrentVersion' in pyproject.toml"
        }
    } else {
        throw "Could not find 'version = \`"...\`"' in pyproject.toml"
    }
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

