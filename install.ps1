param(
    [string]$Destination = "$env:USERPROFILE\.joycode\skills\docx-spec-html"
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourceDir = Join-Path $RepoRoot "docx-spec-html"

if (-not (Test-Path (Join-Path $SourceDir "SKILL.md") -PathType Leaf)) {
    throw "Complete Skill source not found at $SourceDir"
}

New-Item -ItemType Directory -Force -Path $Destination | Out-Null
$null = & robocopy $SourceDir $Destination /MIR /XD ".venv" "__pycache__" /XF "*.pyc" /NFL /NDL /NJH /NJS /NP
if ($LASTEXITCODE -ge 8) {
    throw "Failed to copy the complete Skill (robocopy exit code $LASTEXITCODE)."
}

$UvCommand = Get-Command uv -ErrorAction SilentlyContinue
$UvBin = if ($UvCommand) { $UvCommand.Source } else { "$env:USERPROFILE\.local\bin\uv.exe" }
if (-not (Test-Path $UvBin -PathType Leaf)) {
    Write-Host "Installing uv and managed Python runtime..."
    $PreviousUvInstallDir = $env:UV_INSTALL_DIR
    $env:UV_INSTALL_DIR = "$env:USERPROFILE\.local\bin"
    try {
        Invoke-RestMethod https://astral.sh/uv/install.ps1 | Invoke-Expression
    } finally {
        $env:UV_INSTALL_DIR = $PreviousUvInstallDir
    }
}
if (-not (Test-Path $UvBin -PathType Leaf)) {
    $UvCommand = Get-Command uv -ErrorAction SilentlyContinue
    if ($UvCommand) {
        $UvBin = $UvCommand.Source
    } else {
        throw "uv installation completed but the executable was not found."
    }
}

& $UvBin venv --clear --python 3.12 (Join-Path $Destination ".venv")
if ($LASTEXITCODE -ne 0) {
    throw "Failed to create the Skill Python environment."
}
$SkillPython = Join-Path $Destination ".venv\Scripts\python.exe"
& $UvBin pip install --python $SkillPython -r (Join-Path $Destination "requirements.txt")
if ($LASTEXITCODE -ne 0) {
    throw "Failed to install the Skill Python dependencies."
}

$OfficeCommand = Get-Command officecli -ErrorAction SilentlyContinue
$OfficeCliBin = if ($OfficeCommand) {
    $OfficeCommand.Source
} else {
    "$env:LOCALAPPDATA\OfficeCLI\officecli.exe"
}
if (-not (Test-Path $OfficeCliBin -PathType Leaf)) {
    Write-Host "Installing OfficeCLI from its official distribution..."
    Invoke-RestMethod https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | Invoke-Expression
}
if (-not (Test-Path $OfficeCliBin -PathType Leaf)) {
    $OfficeCommand = Get-Command officecli -ErrorAction SilentlyContinue
    if ($OfficeCommand) {
        $OfficeCliBin = $OfficeCommand.Source
    } else {
        throw "OfficeCLI installation completed but the executable was not found."
    }
}

$env:OFFICECLI_BIN = $OfficeCliBin
& $SkillPython (Join-Path $Destination "scripts\preflight.py")
if ($LASTEXITCODE -ne 0) {
    throw "docx-spec-html preflight failed."
}

Write-Host ""
Write-Host "Installed docx-spec-html at $Destination"
Write-Host "Restart JoyCode so it reloads the Skill."
