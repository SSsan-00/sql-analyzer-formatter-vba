param(
    [string]$BootstrapPath = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
if ([string]::IsNullOrWhiteSpace($BootstrapPath)) {
    $BootstrapPath = & (Join-Path $PSScriptRoot 'build-bootstrap.ps1')
}

$testRoot = Join-Path $env:TEMP ('SqlAnalysisFormatter_Bootstrap_' + [guid]::NewGuid().ToString('N'))
$targetDirectory = Join-Path $testRoot 'out'

try {
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
    powershell -ExecutionPolicy Bypass -File $BootstrapPath -TargetDirectory $targetDirectory | Out-Null

    $workbookPath = Join-Path $targetDirectory 'SqlAnalysisFormatter.xlsm'
    $parserPath = Join-Path $targetDirectory 'SqlAnalysisFormatter.Parser.exe'
    $readmePath = Join-Path $targetDirectory 'README.md'

    foreach ($path in @($workbookPath, $parserPath, $readmePath)) {
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Bootstrap artifact was not created: $path"
        }
    }

    $bootstrapFiles = Get-ChildItem -LiteralPath $targetDirectory -Filter '*bootstrap*.ps1' -File -ErrorAction SilentlyContinue
    if ($bootstrapFiles.Count -gt 0) {
        throw 'Bootstrap source must not be included in artifacts.'
    }

    $versionText = & $parserPath --version
    if ($versionText -notlike 'SqlAnalysisFormatter.Parser*') {
        throw "Parser exe did not start: $versionText"
    }

    Write-Output 'Bootstrap test passed.'
} finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}
