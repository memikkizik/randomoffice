param(
    [ValidateSet('Core', 'Browser', 'All')]
    [string]$Scope = 'All'
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$projectPython = Join-Path $projectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $projectPython)) {
    throw 'Local Python environment missing. Run: uv sync --extra test --extra browser'
}

Push-Location -LiteralPath $projectRoot
try {
    if ($Scope -ne 'Core') {
        & $projectPython -c 'import playwright.sync_api'
        if ($LASTEXITCODE -ne 0) {
            throw 'Browser checks require Playwright. Run: uv sync --extra test --extra browser'
        }
    }

    $nodeCommand = Get-Command node -ErrorAction SilentlyContinue
    if ($nodeCommand) {
        foreach ($scriptPath in @('randomoffice/static/editor.js', 'randomoffice/static/formulas.js','randomoffice/static/workspace.js')) {
            & $nodeCommand.Source --check $scriptPath
            if ($LASTEXITCODE -ne 0) { throw "JavaScript syntax check failed: $scriptPath" }
        }
    } else {
        Write-Warning 'Node.js missing; separate JavaScript syntax checks were not run.'
    }

    & $projectPython -m compileall -q randomoffice tests
    if ($LASTEXITCODE -ne 0) { throw 'Python compile check failed.' }

    & $projectPython -m randomoffice --help | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'CLI help check failed.' }

    $pytestArguments = @('-m', 'pytest', '-o', 'addopts=', '-q')
    if ($Scope -eq 'Core') { $pytestArguments += '--ignore=tests/test_editor_browser.py' }
    if ($Scope -eq 'Browser') { $pytestArguments += 'tests/test_editor_browser.py' }
    & $projectPython @pytestArguments
    if ($LASTEXITCODE -ne 0) { throw "Verification failed for scope: $Scope" }
} finally {
    Pop-Location
}
