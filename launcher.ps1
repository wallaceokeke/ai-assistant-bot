# launcher.ps1
$venvPath = "C:\AI_Assistant\.venv\Scripts\python.exe"
$scriptPath = "C:\AI_Assistant\ai_helper.py"

if (-Not (Test-Path $venvPath)) {
    Write-Host "❌ Virtual environment not found. Please run: python -m venv .venv"
    exit
}

Write-Host "🚀 Launching AI Assistant..."
& $venvPath $scriptPath
