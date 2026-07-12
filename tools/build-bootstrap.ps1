param(
    [string]$WorkbookPath = '',
    [string]$ParserExePath = '',
    [string]$ReadmePath = '',
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot 'SqlAnalysisFormatter.xlsm'
}
if ([string]::IsNullOrWhiteSpace($ParserExePath)) {
    $ParserExePath = Join-Path $repoRoot 'dist\parser\SqlAnalysisFormatter.Parser.exe'
}
if ([string]::IsNullOrWhiteSpace($ReadmePath)) {
    $ReadmePath = Join-Path $repoRoot 'README.md'
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $repoRoot 'dist\bootstrap\SqlAnalysisFormatter.bootstrap.ps1'
}

if (-not (Test-Path -LiteralPath $ParserExePath)) {
    & (Join-Path $PSScriptRoot 'publish-parser.ps1') | Out-Null
}

function ConvertTo-GzipBase64 {
    param([byte[]]$Bytes)

    $output = [System.IO.MemoryStream]::new()
    $gzip = [System.IO.Compression.GZipStream]::new($output, [System.IO.Compression.CompressionLevel]::Optimal, $true)
    $gzip.Write($Bytes, 0, $Bytes.Length)
    $gzip.Dispose()
    try {
        [Convert]::ToBase64String($output.ToArray())
    } finally {
        $output.Dispose()
    }
}

function Split-Base64Literal {
    param([string]$Value)

    $chunkSize = 120
    $chunks = New-Object System.Collections.Generic.List[string]
    for ($index = 0; $index -lt $Value.Length; $index += $chunkSize) {
        $length = [Math]::Min($chunkSize, $Value.Length - $index)
        $chunks.Add($Value.Substring($index, $length))
    }

    ($chunks | ForEach-Object { "    '$_'" }) -join ",`r`n"
}

$workbookBase64 = Split-Base64Literal (ConvertTo-GzipBase64 ([IO.File]::ReadAllBytes((Resolve-Path $WorkbookPath))))
$parserBase64 = Split-Base64Literal (ConvertTo-GzipBase64 ([IO.File]::ReadAllBytes((Resolve-Path $ParserExePath))))
$readmeText = [IO.File]::ReadAllText((Resolve-Path $ReadmePath), [Text.UTF8Encoding]::new($false))

$bootstrap = @"
param(
    [string]`$TargetDirectory = ''
)

`$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace(`$TargetDirectory)) {
    `$TargetDirectory = Join-Path (Get-Location) 'SqlAnalysisFormatter'
}

function Expand-GzipBase64ToFile {
    param(
        [string[]]`$Chunks,
        [string]`$Path
    )

    `$compressedBytes = [Convert]::FromBase64String((`$Chunks -join ''))
    `$inputStream = [System.IO.MemoryStream]::new(`$compressedBytes)
    `$gzip = [System.IO.Compression.GZipStream]::new(`$inputStream, [System.IO.Compression.CompressionMode]::Decompress)
    `$outputStream = [System.IO.File]::Create(`$Path)
    try {
        `$gzip.CopyTo(`$outputStream)
    } finally {
        `$outputStream.Dispose()
        `$gzip.Dispose()
        `$inputStream.Dispose()
    }
}

New-Item -ItemType Directory -Path `$TargetDirectory -Force | Out-Null

`$workbookChunks = @(
$workbookBase64
)

`$parserChunks = @(
$parserBase64
)

`$readmeText = @'
$readmeText
'@

Expand-GzipBase64ToFile `$workbookChunks (Join-Path `$TargetDirectory 'SqlAnalysisFormatter.xlsm')
Expand-GzipBase64ToFile `$parserChunks (Join-Path `$TargetDirectory 'SqlAnalysisFormatter.Parser.exe')
[IO.File]::WriteAllText((Join-Path `$TargetDirectory 'README.md'), `$readmeText, [Text.UTF8Encoding]::new(`$false))

Write-Output "Created: `$TargetDirectory"
"@

New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
[IO.File]::WriteAllText($OutputPath, $bootstrap, [Text.UTF8Encoding]::new($false))
Write-Output $OutputPath
