param(
    [string]$Configuration = 'Release',
    [string]$RuntimeIdentifier = 'win-x64',
    [string]$OutputDirectory = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$projectPath = Join-Path $repoRoot 'tools\SqlAnalysisFormatter.Parser\SqlAnalysisFormatter.Parser.csproj'

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repoRoot 'dist\parser'
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

# 利用者の.NET導入を不要にするため、単一exeへまとめる
dotnet publish $projectPath `
    -c $Configuration `
    -r $RuntimeIdentifier `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=None `
    -p:DebugSymbols=false `
    -o $OutputDirectory

$exePath = Join-Path $OutputDirectory 'SqlAnalysisFormatter.Parser.exe'
if (-not (Test-Path -LiteralPath $exePath)) {
    throw "Parser exe was not created: $exePath"
}

Write-Output $exePath
