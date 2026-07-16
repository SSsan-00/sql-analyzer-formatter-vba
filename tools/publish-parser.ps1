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
# 圧縮単一EXEは毎回の展開コストが大きいため、起動時間を優先する
& dotnet publish $projectPath -c $Configuration -r $RuntimeIdentifier --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=false -p:DebugType=None -p:DebugSymbols=false -o $OutputDirectory
if ($LASTEXITCODE -ne 0) {
    throw "Parser publish failed: exit code $LASTEXITCODE"
}

$exePath = Join-Path $OutputDirectory 'SqlAnalysisFormatter.Parser.exe'
if (-not (Test-Path -LiteralPath $exePath)) {
    throw "Parser exe was not created: $exePath"
}

Write-Output $exePath
