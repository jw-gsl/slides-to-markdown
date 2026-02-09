# PowerPoint Text Extractor Manager (Windows)
# Usage: .\pptx_manager.ps1 <command> [directory] [--keep]

param(
    [Parameter(Position=0)]
    [string]$Command,
    [Parameter(Position=1)]
    [string]$Dir = ".",
    [switch]$Keep
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "pptx_extractor.py"

function Show-Usage {
    Write-Host "Usage: .\pptx_manager.ps1 <command> [directory] [-Keep]"
    Write-Host ""
    Write-Host "Commands:"
    Write-Host "  setup [DIR]         Create folder structure"
    Write-Host "  process [DIR] [-Keep]  Process PPTX files (-Keep leaves originals in input/)"
    Write-Host "  status [DIR]        Show status of folders and files"
    Write-Host "  clean [DIR]         Clean up processed and output folders"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host "  .\pptx_manager.ps1 setup"
    Write-Host "  .\pptx_manager.ps1 process"
    Write-Host "  .\pptx_manager.ps1 process -Keep"
    Write-Host "  .\pptx_manager.ps1 status"
    Write-Host "  .\pptx_manager.ps1 clean"
}

function Get-PythonCommand {
    if (Get-Command python -ErrorAction SilentlyContinue) { return "python" }
    if (Get-Command python3 -ErrorAction SilentlyContinue) { return "python3" }
    if (Get-Command py -ErrorAction SilentlyContinue) { return "py" }
    Write-Host "Error: No Python installation found." -ForegroundColor Red
    Write-Host "Install Python from https://www.python.org/downloads/ and check 'Add to PATH'."
    exit 1
}

function Invoke-Setup {
    $py = Get-PythonCommand
    Write-Host "Setting up folder structure in: $Dir" -ForegroundColor Cyan
    & $py $PythonScript --dir $Dir --setup
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Folder structure created. Place PPTX files in the 'input' folder." -ForegroundColor Green
    } else {
        Write-Host "Failed to create folder structure." -ForegroundColor Red
        exit 1
    }
}

function Invoke-Process {
    $py = Get-PythonCommand
    Write-Host "Processing PPTX files in: $Dir" -ForegroundColor Cyan
    if ($Keep) {
        & $py $PythonScript --dir $Dir --keep
    } else {
        & $py $PythonScript --dir $Dir
    }
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Processing completed." -ForegroundColor Green
    } else {
        Write-Host "Processing failed." -ForegroundColor Red
        exit 1
    }
}

function Show-Status {
    $AbsDir = Resolve-Path $Dir
    Write-Host "Status for: $AbsDir" -ForegroundColor Cyan
    Write-Host ""
    foreach ($folder in @("input", "processed", "output")) {
        $path = Join-Path $AbsDir $folder
        if (Test-Path $path) {
            $count = (Get-ChildItem $path -File | Measure-Object).Count
            Write-Host "  $folder/: $count file(s)" -ForegroundColor Green
        } else {
            Write-Host "  $folder/: missing" -ForegroundColor Yellow
        }
    }
    Write-Host ""
    $inputPath = Join-Path $AbsDir "input"
    if (Test-Path $inputPath) {
        $pptx = Get-ChildItem $inputPath -Filter "*.pptx" -File
        if ($pptx.Count -gt 0) {
            Write-Host "PPTX files ready to process:" -ForegroundColor Yellow
            $pptx | ForEach-Object { Write-Host "  $($_.Name)" }
        } else {
            Write-Host "No PPTX files in input folder." -ForegroundColor Cyan
        }
    }
}

function Invoke-Clean {
    $AbsDir = Resolve-Path $Dir
    $confirm = Read-Host "This will delete all files in 'processed' and 'output'. Continue? (y/N)"
    if ($confirm -eq 'y' -or $confirm -eq 'Y') {
        $processedPath = Join-Path $AbsDir "processed"
        $outputPath = Join-Path $AbsDir "output"
        if (Test-Path $processedPath) { Remove-Item "$processedPath\*" -Force -ErrorAction SilentlyContinue }
        if (Test-Path $outputPath) { Remove-Item "$outputPath\*" -Force -ErrorAction SilentlyContinue }
        Write-Host "Cleaned processed and output folders." -ForegroundColor Green
    } else {
        Write-Host "Cancelled." -ForegroundColor Cyan
    }
}

# Main
switch ($Command) {
    "setup"   { Invoke-Setup }
    "process" { Invoke-Process }
    "status"  { Show-Status }
    "clean"   { Invoke-Clean }
    default   { Show-Usage }
}
