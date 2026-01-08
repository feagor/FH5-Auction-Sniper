param(
    [string]$ReleaseDir = "release",
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $projectRoot
try {
    if (-not $SkipBuild) {
        Write-Host "[build] Running PyInstaller..."
        pyinstaller --noconfirm --onefile --name FH5Sniper main.py
    }

    $exePath = Join-Path $projectRoot 'dist\FH5Sniper.exe'
    if (-not (Test-Path $exePath)) {
        throw "Executable not found at $exePath. Did PyInstaller succeed?"
    }

    $releasePath = Join-Path $projectRoot $ReleaseDir
    if (Test-Path $releasePath) {
        Write-Host "[clean] Removing existing release directory $ReleaseDir"
        Remove-Item $releasePath -Recurse -Force
    }
    New-Item $releasePath -ItemType Directory | Out-Null

    Write-Host "[copy] FH5Sniper.exe -> $ReleaseDir"
    Copy-Item $exePath $releasePath

    $resourceFiles = @(
        'settings.ini',
        'FH5_all_cars_info_v4.xlsx'
    )
    foreach ($file in $resourceFiles) {
        $fullPath = Join-Path $projectRoot $file
        if (Test-Path $fullPath) {
            Write-Host "[copy] $file"
            Copy-Item $fullPath $releasePath
        }
        else {
            Write-Warning "[skip] $file not found"
        }
    }

    $resourceDirs = @('images', 'debug', 'archive')
    foreach ($dir in $resourceDirs) {
        $fullDir = Join-Path $projectRoot $dir
        if (Test-Path $fullDir) {
            Write-Host "[copy] $dir/"
            Copy-Item $fullDir (Join-Path $releasePath $dir) -Recurse
        }
        else {
            Write-Warning "[skip] $dir directory not found"
        }
    }

    Write-Host "[done] Release bundle created at $releasePath"
}
finally {
    Pop-Location
}
