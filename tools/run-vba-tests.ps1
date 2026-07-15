param(
    [string]$ParserExePath = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$workbookPath = Join-Path $repoRoot 'SqlAnalysisFormatter.xlsm'
$mainModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatter.bas'
$testModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatterTests.bas'
$tempWorkbookPath = Join-Path $env:TEMP ('SqlAnalysisFormatter_Tests_' + [guid]::NewGuid().ToString('N') + '.xlsm')
$previousParserExePath = $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
    }
}

Copy-Item -LiteralPath $workbookPath -Destination $tempWorkbookPath -Force

if (-not [string]::IsNullOrWhiteSpace($ParserExePath)) {
    $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = (Resolve-Path $ParserExePath)
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

try {
    $workbook = $excel.Workbooks.Open($tempWorkbookPath)
    $components = $workbook.VBProject.VBComponents

    foreach ($moduleName in @('SqlAnalysisFormatter', 'SqlAnalysisFormatterTests')) {
        try {
            $components.Remove($components.Item($moduleName))
        } catch {
        }
    }

    $components.Import($mainModulePath) | Out-Null
    $components.Import($testModulePath) | Out-Null
    $testResult = [string]$excel.Run("'$tempWorkbookPath'!RunAllSqlAnalysisFormatterTestsForAutomation")
    if ($testResult -ne 'OK') {
        throw "VBA tests failed: $testResult"
    }

    Write-Output 'VBA tests passed.'
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false) | Out-Null
    }
    $excel.Quit()
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Remove-Item -LiteralPath $tempWorkbookPath -Force -ErrorAction SilentlyContinue
    $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = $previousParserExePath
}
