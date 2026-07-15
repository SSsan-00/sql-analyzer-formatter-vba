param(
    [string]$ParserExePath = 'dist\parser\SqlAnalysisFormatter.Parser.exe',
    [string]$CaseId = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$workbookPath = Join-Path $repoRoot 'SqlAnalysisFormatter.xlsm'
$expectationPath = Join-Path $repoRoot 'tests\SqlAnalysisFormatter.OutputExpectations.xlsx'
$fixturePath = Join-Path $repoRoot 'tests\OutputReportCases.json'
$mainModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatter.bas'
$tempWorkbookPath = Join-Path $env:TEMP ('SqlAnalysisFormatter_Golden_' + [guid]::NewGuid().ToString('N') + '.xlsm')
$previousParserExePath = $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE
$env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = (Resolve-Path (Join-Path $repoRoot $ParserExePath))

# COM参照を明示的に解放する
function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
    }
}

# 値が異なる場合はケースとセル番地を付けて停止する
function Assert-Equal {
    param(
        [string]$CaseId,
        [string]$Location,
        [object]$Expected,
        [object]$Actual
    )

    if ([string]$Expected -cne [string]$Actual) {
        throw "$CaseId $Location expected=[$Expected] actual=[$Actual]"
    }
}

# 数値が許容差を超えた場合は停止する
function Assert-Near {
    param(
        [string]$CaseId,
        [string]$Location,
        [double]$Expected,
        [double]$Actual,
        [double]$Tolerance = 0.01
    )

    if ([Math]::Abs($Expected - $Actual) -gt $Tolerance) {
        throw "$CaseId $Location expected=[$Expected] actual=[$Actual]"
    }
}

# 罫線、塗り、フォント、折り返しを期待セルと比較する
function Compare-CellFormat {
    param(
        [string]$CaseId,
        [object]$ExpectedCell,
        [object]$ActualCell
    )

    $address = $ExpectedCell.Address($false, $false)
    Assert-Equal $CaseId "$address fill" $ExpectedCell.Interior.Color $ActualCell.Interior.Color
    Assert-Equal $CaseId "$address font" $ExpectedCell.Font.Name $ActualCell.Font.Name
    Assert-Near $CaseId "$address font-size" $ExpectedCell.Font.Size $ActualCell.Font.Size
    Assert-Equal $CaseId "$address wrap" $ExpectedCell.WrapText $ActualCell.WrapText

    foreach ($borderIndex in @(7, 8, 9, 10)) {
        $expectedBorder = $ExpectedCell.Borders.Item($borderIndex)
        $actualBorder = $ActualCell.Borders.Item($borderIndex)
        Assert-Equal $CaseId "$address border-$borderIndex" $expectedBorder.LineStyle $actualBorder.LineStyle
        if ($expectedBorder.LineStyle -ne -4142) {
            Assert-Equal $CaseId "$address border-$borderIndex-weight" $expectedBorder.Weight $actualBorder.Weight
            Assert-Equal $CaseId "$address border-$borderIndex-color" $expectedBorder.Color $actualBorder.Color
        }
        Release-ComObject $expectedBorder
        Release-ComObject $actualBorder
    }
}

Copy-Item -LiteralPath $workbookPath -Destination $tempWorkbookPath -Force
$fixture = Get-Content -LiteralPath $fixturePath -Encoding UTF8 -Raw | ConvertFrom-Json
$testCases = @($fixture.cases)
if (-not [string]::IsNullOrWhiteSpace($CaseId)) {
    $testCases = @($testCases | Where-Object id -eq $CaseId)
    if ($testCases.Count -eq 0) {
        throw "Output case not found: $CaseId"
    }
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

try {
    $workbook = $excel.Workbooks.Open($tempWorkbookPath)
    $expectationBook = $excel.Workbooks.Open($expectationPath, 0, $true)
    $components = $workbook.VBProject.VBComponents
    try {
        $components.Remove($components.Item('SqlAnalysisFormatter'))
    } catch {
    }
    $components.Import($mainModulePath) | Out-Null

    $definitionSheet = $workbook.Worksheets.Item('変換定義')
    $sqlSheet = $workbook.Worksheets.Item('SQL解析')
    $outputSheet = $workbook.Worksheets.Item('アウトプット')
    $formatColumns = @(1, 6, 7, 15, 17, 18, 19, 31, 32, 36, 37, 90)
    $caseIndex = 0

    foreach ($testCase in $testCases) {
        $caseIndex++
        $excel.Run("'$tempWorkbookPath'!ClearData", $false) | Out-Null

        $definitionRow = 2
        $tableProperties = @($testCase.tables.PSObject.Properties)
        if ($tableProperties.Count -eq 0) {
            $tableProperties = @([pscustomobject]@{ Name = '__dummy__'; Value = '未使用' })
        }
        foreach ($table in $tableProperties) {
            $definitionSheet.Cells.Item($definitionRow, 1).Value2 = [string]$table.Name
            $definitionSheet.Cells.Item($definitionRow, 2).Value2 = [string]$table.Value
            $definitionSheet.Cells.Item($definitionRow, 3).Value2 = '__unused__'
            $definitionSheet.Cells.Item($definitionRow, 4).Value2 = '未使用'
            $definitionRow++
        }

        $sqlRow = 2
        foreach ($line in $testCase.sql_lines) {
            $sqlSheet.Cells.Item($sqlRow, 1).Value2 = [string]$line
            $sqlRow++
        }

        $excel.Run("'$tempWorkbookPath'!AnalyzeQueries", $false) | Out-Null
        $expectedSheet = $expectationBook.Worksheets.Item([string]$testCase.id)
        $expectedRowCount = $expectedSheet.UsedRange.Rows.Count
        $expectedRange = $expectedSheet.Range("A1:CL$expectedRowCount")
        $actualRange = $outputSheet.Range("A1:CL$expectedRowCount")
        $expectedValues = $expectedRange.Value2
        $actualValues = $actualRange.Value2

        for ($row = 1; $row -le $expectedRowCount; $row++) {
            for ($column = 1; $column -le 90; $column++) {
                Assert-Equal ([string]$testCase.id) ("R{0}C{1}" -f $row, $column) $expectedValues[$row, $column] $actualValues[$row, $column]
            }
            Assert-Near ([string]$testCase.id) "row-$row-height" $expectedSheet.Rows.Item($row).RowHeight $outputSheet.Rows.Item($row).RowHeight

            foreach ($column in $formatColumns) {
                $expectedCell = $expectedSheet.Cells.Item($row, $column)
                $actualCell = $outputSheet.Cells.Item($row, $column)
                Compare-CellFormat ([string]$testCase.id) $expectedCell $actualCell
                Release-ComObject $expectedCell
                Release-ComObject $actualCell
            }
        }

        if ($caseIndex -eq 1) {
            for ($column = 1; $column -le 90; $column++) {
                Assert-Near ([string]$testCase.id) "column-$column-width" $expectedSheet.Columns.Item($column).ColumnWidth $outputSheet.Columns.Item($column).ColumnWidth
            }
        }

        $outputSheet.Activate() | Out-Null
        if ($excel.ActiveWindow.DisplayGridlines) {
            throw "$($testCase.id) gridlines should be hidden"
        }

        Release-ComObject $actualRange
        Release-ComObject $expectedRange
        Release-ComObject $expectedSheet
        if ($caseIndex % 10 -eq 0 -or $caseIndex -eq $testCases.Count) {
            Write-Output ("Output golden progress: {0}/{1}" -f $caseIndex, $testCases.Count)
        }
    }

    Write-Output ("Output golden tests passed: {0} cases." -f $testCases.Count)
} finally {
    if ($null -ne $expectationBook) {
        $expectationBook.Close($false) | Out-Null
    }
    if ($null -ne $workbook) {
        $workbook.Close($false) | Out-Null
    }
    $excel.Quit()
    foreach ($comObject in @($outputSheet, $sqlSheet, $definitionSheet, $components, $expectationBook, $workbook, $excel)) {
        Release-ComObject $comObject
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Remove-Item -LiteralPath $tempWorkbookPath -Force -ErrorAction SilentlyContinue
    $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = $previousParserExePath
}
