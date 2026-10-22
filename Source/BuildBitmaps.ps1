<#
.SYNOPSIS
Compiles the listed resource scripts.

.DESCRIPTION
Runs C:\Tools\rc.exe on each listed resource script and writes the
corresponding .res file to the parent directory of this script's directory.
#>

begin {
    $ErrorActionPreference = 'Stop'
    $rcExe = 'C:\Tools\rc.exe'

    $FileList = @(
        'CharacterBuilderLite',
        'Compendium',
        'CannithCrafting'
    )

    if (-not (Test-Path -LiteralPath $rcExe -PathType Leaf)) {
        throw "Resource compiler not found: $rcExe"
    }
}

process {
    foreach ($file in $FileList) {
        write-host "Compiling $file ..." -ForegroundColor Green
        $rcFile = Join-Path $PSScriptRoot "$file.rc"
        $outputFile = Join-Path (Split-Path $PSScriptRoot -Parent) "$file.res"

        if (-not (Test-Path -LiteralPath $rcFile -PathType Leaf)) {
            throw "Resource script not found: $rcFile"
        }

        & $rcExe $rcFile
        if ($LASTEXITCODE -ne 0) {
            throw "RC.EXE failed for $rcFile with exit code $LASTEXITCODE."
        }

        Write-Host "  Created $outputFile" -ForegroundColor Cyan
    }
}