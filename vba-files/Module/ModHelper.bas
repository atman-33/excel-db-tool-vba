Sub CreateSheetIfNotExists(sheetName As String)
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    
    ' シートが存在するか確認する
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' シートが存在しない場合、新しいシートを作成する
    If Not sheetExists Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = sheetName
        Debug.Print "新しいシート '" & sheetName & "' を作成しました。"
    Else
        Debug.Print "指定したシート名 '" & sheetName & "' は既に存在します。"
    End If
End Sub

' Sub TestCreateSheetIfNotExists()
'     CreateSheetIfNotExists "新しいシート"
' End Sub

Function GetTableFromCell(sheetName As String, cellAddress As String) As ListObject
    Dim tbl As ListObject
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        MsgBox "指定したシートが存在しません。"
        Exit Function
    End If
    
    Set tbl = ws.Range(cellAddress).ListObject
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "指定したセルはテーブルに含まれていません。"
        Set GetTableFromCell = Nothing
    Else
        Set GetTableFromCell = tbl
    End If
End Function

' Sub TestGetTableFromCell()
'     Dim tbl As ListObject
'     Dim sheetName As String
'     Dim cellAddress As String
    
'     ' シート名とセルのアドレスを指定
'     sheetName = "1"
'     cellAddress = "A1"
    
'     ' テーブルを取得
'     Set tbl = GetTableFromCell(sheetName, cellAddress)
    
'     If Not tbl Is Nothing Then
'         MsgBox "指定したセルが含まれるテーブルの名前は " & tbl.Name & " です。"
'     End If
' End Sub

