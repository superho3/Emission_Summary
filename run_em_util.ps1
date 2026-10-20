Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptDir

try {
    $venvPython = Join-Path $scriptDir ".venv\Scripts\python.exe"

    if (Test-Path $venvPython) {
        & $venvPython "EM_SUM.py" @args
    }
    else {
        python "EM_SUM.py" @args
    }
}
finally {
    Pop-Location
}