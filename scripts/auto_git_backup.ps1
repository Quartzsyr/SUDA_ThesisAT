param(
    [string]$CommitPrefix = "auto backup"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $repoRoot

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw "git is not installed or not available in PATH."
}

$insideWorkTree = & git rev-parse --is-inside-work-tree 2>$null
if ($LASTEXITCODE -ne 0 -or $insideWorkTree -ne "true") {
    throw "Current directory is not a git repository: $repoRoot"
}

$statusLines = @(& git status --porcelain)
if (-not $statusLines -or $statusLines.Count -eq 0) {
    Write-Host "No changes to back up."
    exit 0
}

& git add -A
if ($LASTEXITCODE -ne 0) {
    throw "git add failed."
}

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
& git commit -m "$CommitPrefix $timestamp"
if ($LASTEXITCODE -ne 0) {
    throw "git commit failed."
}

Write-Host "Backup commit created at $timestamp"
