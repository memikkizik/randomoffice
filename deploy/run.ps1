# Quick deployment launcher for Windows
param (
    [string]$Host = "127.0.0.1",
    [int]$Port = 8000
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Starting RandomOffice Enterprise Service" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# Check if uv or virtual environment exists
if (Test-Path ".venv\Scripts\uvicorn.exe") {
    Write-Host "Starting Uvicorn server on http://$($Host):$($Port)..." -ForegroundColor Green
    & ".\.venv\Scripts\uvicorn.exe" randomoffice.api.app:app --host $Host --port $Port
} else {
    Write-Host "Virtualenv not found, using uv run..." -ForegroundColor Yellow
    uv run uvicorn randomoffice.api.app:app --host $Host --port $Port
}
