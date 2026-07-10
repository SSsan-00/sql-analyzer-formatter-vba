Attribute VB_Name = "SqlAnalyzerFormatterTests"
Option Explicit

'@TestModule
'@Folder("Tests")

Private Const COL_SQL As Long = 1
Private Const COL_RESULT As Long = 2
Private Const COL_REPLACEMENT As Long = 3

' テスト一式をまとめて実行
Public Sub RunAllSqlAnalyzerFormatterTests(Optional ByVal showMessage As Boolean = True)
    On Error GoTo TestFail

    SetupWorkbook_CreatesOutputSheet
    AnalyzeQueries_ConvertsCrudFixtures

    If showMessage Then
        MsgBox "SqlAnalyzerFormatter tests passed.", vbInformation
    End If
    Exit Sub

TestFail:
    If showMessage Then
        MsgBox Err.Description, vbCritical
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'@TestMethod("SetupWorkbook")
Public Sub SetupWorkbook_CreatesOutputSheet()
    SetupWorkbook

    AssertWorksheetExists OutputSheetName()
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_ConvertsCrudFixtures()
    Dim wsSql As Worksheet

    ArrangeCrudFixtures
    AnalyzeQueries False

    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    AssertAnalyzeRow wsSql, 2, ExpectedSelectSql(), _
        Array("users." & FullNameText(), "users." & UserIdText(), "orders." & OrderIdText(), StatusText(), "orders." & OrderUserIdText())
    AssertAnalyzeRow wsSql, 3, ExpectedInsertSql(), _
        Array(StatusText(), CreatedAtText(), "orders." & OrderIdText(), "users." & UserIdText(), "orders." & AmountText())
    AssertAnalyzeRow wsSql, 4, ExpectedUpdateSql(), _
        Array("users." & FullNameText(), UpdatedAtText(), StatusText(), "users." & UserIdText())
    AssertAnalyzeRow wsSql, 5, ExpectedDeleteSql(), _
        Array("orders." & OrderIdText(), "order_items." & DetailOrderIdText(), "order_items." & ProductIdText(), StatusText())
End Sub

' CRUDを含む解析テスト用データを作成
Private Sub ArrangeCrudFixtures()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents

    SeedReferenceDefinitions wsRef
    SeedCrudQueries wsSql
End Sub

' CRUDサンプルで使う変換定義を投入
Private Sub SeedReferenceDefinitions(ByVal ws As Worksheet)
    PutDefinition ws, 2, "users", UserTableText(), "user_id", UserIdText()
    PutDefinition ws, 3, "users", UserTableText(), "name", FullNameText()
    PutDefinition ws, 4, "users", UserTableText(), "email", MailText()
    PutDefinition ws, 5, "orders", OrderTableText(), "order_id", OrderIdText()
    PutDefinition ws, 6, "orders", OrderTableText(), "user_id", OrderUserIdText()
    PutDefinition ws, 7, "orders", OrderTableText(), "amount", AmountText()
    PutDefinition ws, 8, "order_items", OrderItemTableText(), "order_id", DetailOrderIdText()
    PutDefinition ws, 9, "order_items", OrderItemTableText(), "product_id", ProductIdText()
    PutDefinition ws, 10, "order_items", OrderItemTableText(), "quantity", QuantityText()
    PutDefinition ws, 11, "-", "", "status", StatusText()
    PutDefinition ws, 12, "-", "", "created_at", CreatedAtText()
    PutDefinition ws, 13, "-", "", "updated_at", UpdatedAtText()
End Sub

' CRUDサンプルクエリを投入
Private Sub SeedCrudQueries(ByVal ws As Worksheet)
    ws.Cells(2, COL_SQL).Value = "select trim(users.name) as name, users.user_id, orders.order_id, status from users inner join orders on users.user_id = orders.user_id where status = 'ACTIVE'"
    ws.Cells(3, COL_SQL).Value = "insert into orders (order_id, user_id, amount, status, created_at) select orders.order_id, users.user_id, orders.amount, status, created_at from users"
    ws.Cells(4, COL_SQL).Value = "update users set users.name = 'Taro', updated_at = CURRENT_TIMESTAMP, status = 'ACTIVE' where users.user_id = :user_id"
    ws.Cells(5, COL_SQL).Value = "delete from orders where orders.order_id in (select order_items.order_id from order_items where order_items.product_id = :product_id) and status = 'CANCELLED'"
End Sub

Private Function ExpectedSelectSql() As String
    ExpectedSelectSql = "select trim(users." & FullNameText() & ") as name, users." & UserIdText() & _
        ", orders." & OrderIdText() & ", " & StatusText() & _
        " from users inner join orders on users." & UserIdText() & " = orders." & OrderUserIdText() & _
        " where " & StatusText() & " = 'ACTIVE'"
End Function

Private Function ExpectedInsertSql() As String
    ExpectedInsertSql = "insert into orders (order_id, user_id, amount, " & StatusText() & ", " & CreatedAtText() & _
        ") select orders." & OrderIdText() & ", users." & UserIdText() & ", orders." & AmountText() & _
        ", " & StatusText() & ", " & CreatedAtText() & " from users"
End Function

Private Function ExpectedUpdateSql() As String
    ExpectedUpdateSql = "update users set users." & FullNameText() & " = 'Taro', " & _
        UpdatedAtText() & " = CURRENT_TIMESTAMP, " & StatusText() & _
        " = 'ACTIVE' where users." & UserIdText() & " = :user_id"
End Function

Private Function ExpectedDeleteSql() As String
    ExpectedDeleteSql = "delete from orders where orders." & OrderIdText() & _
        " in (select order_items." & DetailOrderIdText() & _
        " from order_items where order_items." & ProductIdText() & _
        " = :product_id) and " & StatusText() & " = 'CANCELLED'"
End Function

Private Sub PutDefinition(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal tableId As String, ByVal tableName As String, ByVal fieldId As String, ByVal fieldName As String)
    ws.Cells(rowNumber, 1).Value = tableId
    ws.Cells(rowNumber, 2).Value = tableName
    ws.Cells(rowNumber, 3).Value = fieldId
    ws.Cells(rowNumber, 4).Value = fieldName
End Sub

Private Sub AssertAnalyzeRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal expectedSql As String, ByVal expectedReplacements As Variant)
    Dim index As Long
    Dim replacementCount As Long

    AssertCellValue ws.Cells(rowNumber, COL_RESULT), expectedSql
    For index = LBound(expectedReplacements) To UBound(expectedReplacements)
        AssertCellValue ws.Cells(rowNumber, COL_REPLACEMENT + index), CStr(expectedReplacements(index))
    Next index

    replacementCount = UBound(expectedReplacements) - LBound(expectedReplacements) + 1
    AssertCellValue ws.Cells(rowNumber, COL_REPLACEMENT + replacementCount), ""
End Sub

Private Sub AssertWorksheetExists(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Fail "Worksheet not found: " & sheetName
    End If
End Sub

Private Sub AssertCellValue(ByVal cell As Range, ByVal expected As String)
    Dim actual As String

    actual = CStr(cell.Value)
    If actual <> expected Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " expected=[" & expected & "] actual=[" & actual & "]"
    End If
End Sub

Private Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 513, "SqlAnalyzerFormatterTests", message
End Sub

' VBEインポート時の文字化けを避けるため、テスト文字列もコードポイントで保持
Private Function W(ParamArray codes() As Variant) As String
    Dim index As Long
    Dim resultText As String

    For index = LBound(codes) To UBound(codes)
        resultText = resultText & ChrW$(CLng(codes(index)))
    Next index

    W = resultText
End Function

Private Function ReferenceSheetName() As String
    ReferenceSheetName = W(&H5909, &H63DB, &H5B9A, &H7FA9)
End Function

Private Function SqlSheetName() As String
    SqlSheetName = "SQL" & W(&H89E3, &H6790)
End Function

Private Function OutputSheetName() As String
    OutputSheetName = W(&H30A2, &H30A6, &H30C8, &H30D7, &H30C3, &H30C8)
End Function

Private Function UserTableText() As String
    UserTableText = W(&H30E6, &H30FC, &H30B6, &H30FC)
End Function

Private Function OrderTableText() As String
    OrderTableText = W(&H6CE8, &H6587)
End Function

Private Function OrderItemTableText() As String
    OrderItemTableText = W(&H6CE8, &H6587, &H660E, &H7D30)
End Function

Private Function UserIdText() As String
    UserIdText = UserTableText() & "ID"
End Function

Private Function FullNameText() As String
    FullNameText = W(&H6C0F, &H540D)
End Function

Private Function MailText() As String
    MailText = W(&H30E1, &H30FC, &H30EB)
End Function

Private Function OrderIdText() As String
    OrderIdText = OrderTableText() & "ID"
End Function

Private Function OrderUserIdText() As String
    OrderUserIdText = OrderTableText() & UserTableText() & "ID"
End Function

Private Function AmountText() As String
    AmountText = W(&H91D1, &H984D)
End Function

Private Function DetailOrderIdText() As String
    DetailOrderIdText = W(&H660E, &H7D30, &H6CE8, &H6587) & "ID"
End Function

Private Function ProductIdText() As String
    ProductIdText = W(&H5546, &H54C1) & "ID"
End Function

Private Function QuantityText() As String
    QuantityText = W(&H6570, &H91CF)
End Function

Private Function StatusText() As String
    StatusText = W(&H72B6, &H614B)
End Function

Private Function CreatedAtText() As String
    CreatedAtText = W(&H4F5C, &H6210, &H65E5, &H6642)
End Function

Private Function UpdatedAtText() As String
    UpdatedAtText = W(&H66F4, &H65B0, &H65E5, &H6642)
End Function
