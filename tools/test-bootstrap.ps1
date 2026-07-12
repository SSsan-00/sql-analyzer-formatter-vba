param(
    [string]$UserBootstrapPath = '',
    [string]$DeveloperBootstrapPath = ''
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($UserBootstrapPath)) {
    $UserBootstrapPath = & (Join-Path $PSScriptRoot 'build-bootstrap.ps1') -Audience User
}
if ([string]::IsNullOrWhiteSpace($DeveloperBootstrapPath)) {
    $DeveloperBootstrapPath = & (Join-Path $PSScriptRoot 'build-bootstrap.ps1') -Audience Developer
}

function Assert-PathExists {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Bootstrap artifact was not created: $Path"
    }
}

function Assert-PathMissing {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        throw "Unexpected bootstrap artifact was created: $Path"
    }
}

function Assert-NoBootstrapSource {
    param([string]$TargetDirectory)

    $bootstrapFiles = Get-ChildItem -LiteralPath $TargetDirectory -Filter '*bootstrap*.ps1' -Recurse -File -ErrorAction SilentlyContinue
    if ($bootstrapFiles.Count -gt 0) {
        throw 'Bootstrap source must not be included in artifacts.'
    }
}

function Invoke-Bootstrap {
    param(
        [string]$BootstrapPath,
        [string]$TargetDirectory
    )

    powershell -ExecutionPolicy Bypass -File $BootstrapPath -TargetDirectory $TargetDirectory | Out-Null
}

function Assert-ParserStarts {
    param([string]$ParserPath)

    $versionText = & $ParserPath --version
    if ($versionText -notlike 'SqlAnalysisFormatter.Parser*') {
        throw "Parser exe did not start: $versionText"
    }
}

$testRoot = Join-Path $env:TEMP ('SqlAnalysisFormatter_Bootstrap_' + [guid]::NewGuid().ToString('N'))
$userTarget = Join-Path $testRoot 'user'
$developerTarget = Join-Path $testRoot 'developer'

try {
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

    Invoke-Bootstrap $UserBootstrapPath $userTarget
    Assert-PathExists (Join-Path $userTarget 'SqlAnalysisFormatter.bas')
    Assert-PathExists (Join-Path $userTarget 'SqlAnalysisFormatter.Parser.exe')
    Assert-PathExists (Join-Path $userTarget 'README.md')
    Assert-PathMissing (Join-Path $userTarget 'SqlAnalysisFormatter.xlsm')
    Assert-PathMissing (Join-Path $userTarget 'src')
    Assert-PathMissing (Join-Path $userTarget 'tests')
    Assert-PathMissing (Join-Path $userTarget 'tools')
    Assert-NoBootstrapSource $userTarget
    Assert-ParserStarts (Join-Path $userTarget 'SqlAnalysisFormatter.Parser.exe')

    Invoke-Bootstrap $DeveloperBootstrapPath $developerTarget
    Assert-PathExists (Join-Path $developerTarget 'SqlAnalysisFormatter.bas')
    Assert-PathExists (Join-Path $developerTarget 'SqlAnalysisFormatter.Parser.exe')
    Assert-PathExists (Join-Path $developerTarget 'README.md')
    Assert-PathExists (Join-Path $developerTarget 'docs\DEVELOPER_GUIDE.md')
    Assert-PathExists (Join-Path $developerTarget 'src\vba\SqlAnalysisFormatter.bas')
    Assert-PathExists (Join-Path $developerTarget 'src\vba\SqlAnalysisFormatterTests.bas')
    Assert-PathExists (Join-Path $developerTarget 'tools\SqlAnalysisFormatter.Parser\TsqlAstParser.cs')
    Assert-PathExists (Join-Path $developerTarget 'tests\SqlAnalysisFormatter.Parser.Tests\TsqlAstParserTests.cs')
    Assert-PathExists (Join-Path $developerTarget 'tests\CRUD_TEST_CASES.md')
    Assert-PathExists (Join-Path $developerTarget 'SqlAnalysisFormatter.sln')
    Assert-PathMissing (Join-Path $developerTarget 'SqlAnalysisFormatter.xlsm')
    Assert-NoBootstrapSource $developerTarget
    Assert-ParserStarts (Join-Path $developerTarget 'SqlAnalysisFormatter.Parser.exe')

    Write-Output 'Bootstrap test passed.'
} finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}
