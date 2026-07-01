#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

function Test-RealPython {
    param(
        [string]$Executable,
        [string[]]$PrefixArgs = @()
    )

    if (-not (Get-Command $Executable -ErrorAction SilentlyContinue)) {
        return $false
    }

    $command = Get-Command $Executable
    if ($command.Source -like "*\WindowsApps\*") {
        return $false
    }

    try {
        $output = & $Executable @($PrefixArgs + @("--version")) 2>&1
        return ($output -match "Python \d")
    } catch {
        return $false
    }
}

function Find-PythonLauncher {
    if (Test-RealPython -Executable "py" -PrefixArgs @("-3")) {
        return @{
            Executable = "py"
            PrefixArgs = @("-3")
        }
    }

    foreach ($candidate in @("python", "python3")) {
        if (Test-RealPython -Executable $candidate) {
            return @{
                Executable = $candidate
                PrefixArgs = @()
            }
        }
    }

    return $null
}

function Write-PythonInstallHelp {
    Write-Host ""
    Write-Host "[ERROR] Python 3.10+ is required." -ForegroundColor Red
    Write-Host "Install from https://www.python.org/downloads/windows/"
    Write-Host "Or run: winget install Python.Python.3.12"
    Write-Host "Disable Microsoft Store python aliases if needed."
    Write-Host ""
}

$launcher = Find-PythonLauncher
if (-not $launcher) {
    Write-PythonInstallHelp
    exit 1
}

$PythonExe = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path $PythonExe)) {
    Write-Host "Creating virtual environment..."
    & $launcher.Executable @($launcher.PrefixArgs + @("-m", "venv", ".venv"))
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $PythonExe)) {
        Write-Host "[ERROR] Failed to create virtual environment." -ForegroundColor Red
        Write-PythonInstallHelp
        exit 1
    }
}

$Python = $PythonExe
$Pip = Join-Path $Root ".venv\Scripts\pip.exe"

& $Python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] pip upgrade failed." -ForegroundColor Red
    exit 1
}

& $Pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] dependency install failed." -ForegroundColor Red
    exit 1
}

$iconPath = Join-Path $Root "assets\mailbatch.ico"
if (-not (Test-Path $iconPath)) {
    Write-Host "Generating application icon..."
    & $Python (Join-Path $Root "scripts\generate_icon.py")
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] icon generation failed." -ForegroundColor Red
        exit 1
    }
}

$iconArg = @()
$dataArg = @()
if (Test-Path $iconPath) {
    $iconArg = @("--icon", $iconPath)
    $dataArg = @("--add-data", "$iconPath;assets")
}

$appHtml = Join-Path $Root "app.html"
$appHtmlArg = @()
if (Test-Path $appHtml) {
    $appHtmlArg = @("--add-data", "$appHtml;.")
} else {
    Write-Host "[ERROR] app.html not found." -ForegroundColor Red
    exit 1
}

& $Python -m PyInstaller --noconfirm --windowed --onefile --name MailBatch `
    --hidden-import webview `
    --hidden-import webview.platforms.winforms `
    @iconArg @dataArg @appHtmlArg app.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] PyInstaller build failed." -ForegroundColor Red
    exit 1
}

$ExePath = Join-Path $Root "dist\MailBatch.exe"
if (-not (Test-Path $ExePath)) {
    Write-Host "[ERROR] Executable not found: $ExePath" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Build complete: $ExePath" -ForegroundColor Green
