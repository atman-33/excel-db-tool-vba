Attribute VB_Name = "ModCommon"
Dim tabInfo As New ClsTable
Public dbinfo As New ClsDBInfo

Public Sub btnReplaceFiles_Click()
    Dim thisSHeet As Worksheet
    Set thisSHeet = ActiveSheet
    Call ReplaceFiles(thisSHeet.Range("C1"), _
                        thisSHeet.Range("C2"), _
                        thisSHeet.Range("C3"), _
                        thisSHeet.Range("C4") _
                        )
End Sub


Public Sub btnQuery_Click()
    tabInfo.ShowMetaTable ActiveSheet
    tabInfo.ShowDataRows ActiveSheet
End Sub

Public Sub btnBackupData_Click()
    Dim thisSHeet As Worksheet
    Set thisSHeet = ActiveSheet
    
    Dim newSheet As Worksheet
    Set newSheet = ThisWorkbook.Sheets.Add(, thisSHeet)
    newSheet.Name = thisSHeet.Name & "_" & Format(Now, "HHmmss")
    
    thisSHeet.Activate
    thisSHeet.Range("A2:H3").Copy
    newSheet.Activate
    newSheet.Range("A1").Select
    newSheet.Paste
    Application.CutCopyMode = False
    
    thisSHeet.rows("7:" & thisSHeet.UsedRange.rows.count).Copy
    newSheet.Range("A3").Select
    ActiveSheet.Paste
    
End Sub


Public Sub btnUpdate_Click()
    tabInfo.ShowMetaTable ActiveSheet
    tabInfo.UpdateData ActiveSheet
    
    tabInfo.ShowDataRows ActiveSheet
End Sub
Public Sub btnMakeSQL_Click()
    
    Call MakeBatchSQL(ActiveSheet)
    
End Sub
Public Sub btnExecSQL_Click()
    Dim batchSQLList As New Collection
    Call FillBatchSQL(ActiveSheet, batchSQLList)
    Call ShowBatchResult(Sheet6, batchSQLList)
End Sub

Private Sub MakeBatchSQL(thisSHeet As Worksheet)
    Dim thisRange As Range
    Set thisRange = findrange(thisSHeet.Range("A:A"), "#SQL")
    If thisRange Is Nothing Then Exit Sub
    
    Dim r1 As Integer
    'SQL開始行
    r1 = thisRange.row + 2
    
    Dim rowSql As String
    Dim sqlinfo As ClsSQL
    Dim isSelect As Boolean
    '空まで、繰り返し
    Do While Not IsEmpty(thisSHeet.Range("C" & r1))
        rowSql = ""
        isSelect = False
        'SELECT + FROM + WHERE + GROUPBY + ORDERBY
        Select Case UCase(thisSHeet.Range("C" & r1))
        Case "SELECT"
            isSelect = True
            rowSql = rowSql & thisSHeet.Range("C" & r1) & " " _
                        & thisSHeet.Range("D" & r1) _
                        & " FROM " _
                        & thisSHeet.Range("E" & r1)

        Case "UPDATE"
            rowSql = rowSql & thisSHeet.Range("C" & r1) & " " _
                        & thisSHeet.Range("E" & r1) _
                        & " SET " _
                        & thisSHeet.Range("D" & r1)
                        
        Case "INSERT"
            rowSql = rowSql & thisSHeet.Range("C" & r1) & " " _
                        & thisSHeet.Range("E" & r1) _
                        & " ( " _
                        & thisSHeet.Range("D" & r1) _
                        & " )VALUES( " _
                        & thisSHeet.Range("I" & r1) _
                        & " ); "
                        
        Case "DELETE"
            rowSql = rowSql & thisSHeet.Range("C" & r1) & " " _
                        & " FROM " _
                        & thisSHeet.Range("E" & r1)
        Case ""
            Exit Do
        Case Else
            MsgBox "エラー：SELECT,INSERT,UPDATE,DELETE以外は入力禁止", vbCritical
            End
        End Select
        
        If Not IsEmpty(thisSHeet.Range("F" & r1)) Then
            rowSql = rowSql & " WHERE " & thisSHeet.Range("F" & r1)
        End If
        If Not IsEmpty(thisSHeet.Range("G" & r1)) Then
            rowSql = rowSql & " GROUP BY " & thisSHeet.Range("G" & r1)
        End If
        If Not IsEmpty(thisSHeet.Range("H" & r1)) Then
            rowSql = rowSql & " ORDER BY " & thisSHeet.Range("H" & r1)
        End If
        If isSelect And Not IsEmpty(thisSHeet.Range("J" & r1)) Then
            rowSql = " SELECT * FROM (" & rowSql & ") WHERE ROWNUM<=" & thisSHeet.Range("J" & r1)
        End If
        thisSHeet.Range("B" & r1) = rowSql
        
        '=========================
        r1 = r1 + 1
        DoEvents
    Loop
    
End Sub

Private Sub FillBatchSQL(thisSHeet As Worksheet, batchSQLList As Collection)
    Dim thisRange As Range
    Set thisRange = findrange(thisSHeet.Range("A:A"), "#SQL")
    If thisRange Is Nothing Then Exit Sub
    
    Dim r1 As Integer
    'SQL開始行
    r1 = thisRange.row + 2
    
    Dim strSQL As String
    Dim sqlinfo As ClsSQL
    
    '空まで、繰り返し
    Do While Not IsEmpty(thisSHeet.Range("B" & r1))
        If strSQL <> "" Then
            strSQL = strSQL & ";"
        End If
        
        Set sqlinfo = New ClsSQL
        sqlinfo.id = thisSHeet.Range("A" & r1)
        sqlinfo.sql = thisSHeet.Range("B" & r1)
        sqlinfo.isSelect = IsTestOK(sqlinfo.sql, "\s*SELECT")
        batchSQLList.Add sqlinfo
        strSQL = strSQL & thisSHeet.Range("B" & r1)
        
        '=========================
        r1 = r1 + 1
        DoEvents
    Loop
    strSQL = Replace(strSQL, vbLf, " ")
    'DB接続情報取得
    dbinfo.Init thisSHeet
    
    Dim results As String
    results = dbinfo.Batch(strSQL)
    Dim lst() As String
    lst = Split(results, ";")
    Dim i As Integer
    For i = 0 To UBound(lst)
        If IsTestOK(lst(i), "error") Then
            batchSQLList(i + 1).ErrMsg = lst(i)
        Else
            batchSQLList(i + 1).Result = lst(i)
        End If
    Next
End Sub

Private Sub ShowBatchResult(thisSHeet As Worksheet, batchSQLList As Collection)
    With thisSHeet.Cells
        .Clear
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .NumberFormatLocal = "@"
    End With
    thisSHeet.Activate
    
    Dim r As Integer
    r = 2
    Dim sqlinfo As ClsSQL
    
    Dim x As Integer
    Dim r1 As Integer
    Dim dt As New ClsDataTable
    For Each sqlinfo In batchSQLList
        thisSHeet.Cells(r, 1) = sqlinfo.id
        thisSHeet.Cells(r, 2) = sqlinfo.sql
        thisSHeet.Cells(r, 2).WrapText = False
        thisSHeet.Range(r & ":" & r).Interior.colorindex = 24
        r = r + 1
        If sqlinfo.ErrMsg <> "" Then
           thisSHeet.Cells(r, 2) = sqlinfo.ErrMsg
           thisSHeet.Cells(r, 2).Interior.colorindex = 3
        Else
            'データ詳細を表示
            If sqlinfo.isSelect Then
                dt.Fill sqlinfo.Result
                thisSHeet.Cells(r, 1) = "検索件数:" & dt.DataRows.count
                thisSHeet.Cells(r, 1).Interior.colorindex = 4
                For x = 1 To dt.ColumnNames.count
                    thisSHeet.Cells(r, x + 1) = dt.ColumnNames(x)
                Next
                ShowHeaderRangeStyle thisSHeet, r, 2, r, dt.ColumnNames.count + 1
                r1 = r + 1
                For y = 1 To dt.DataRows.count
                    r = r + 1
                    On Error Resume Next
                    For x = 1 To dt.ColumnNames.count
                        thisSHeet.Cells(r, x + 1) = dt.DataRows(y)(x)
                        If Err.Number <> 0 Then
                            Exit For
                        End If
                    Next
                    If Err.Number <> 0 Then
                        sqlinfo.ErrMsg = Err.Description
                        Exit For
                    End If
                Next
                r = r + 1
                ShowDataRangeStyle thisSHeet, r1, 2, r - 1, dt.ColumnNames.count + 1
            Else
                thisSHeet.Cells(r, 1) = "更新件数:" & sqlinfo.Result
                thisSHeet.Cells(r, 1).Interior.colorindex = 8
            End If
        End If
        r = r + 1
    Next
End Sub


Public Sub SplitToDictionary(source As String, Map As Dictionary, Optional isTrim As Boolean = False)
    Dim ms() As String
    ms = Split(source, "|")
    Map.RemoveAll
    
    Dim i As Integer
    i = 0
    Dim m
    For Each m In ms
        If isTrim Then
            Map.Add Trim(m), i
        Else
            Map.Add m, i
        End If
        i = i + 1
    Next
End Sub

Public Sub SplitDataToCollection(source As String, Map As Collection, Optional isTrim As Boolean = False)
    Dim ms() As String
    ms = Split(source, "|")
    RemoveAll Map
    
    Dim m
    For Each m In ms
        If isTrim Then
            Map.Add Trim(m)
        Else
            Map.Add m
        End If
    Next
End Sub

Public Function GetMatchCollecion(source As String, pattern As String) As MatchCollection

    Dim reg As New RegExp
    reg.pattern = pattern
    reg.MultiLine = True
    reg.IgnoreCase = True
    reg.Global = True
    Set GetMatchCollecion = reg.Execute(source)

End Function

Public Sub RemoveAll(lst As Collection)
 While lst.count > 0
    lst.Remove 1
 Wend
End Sub

Public Function GetIndex(lst As Collection, v) As Integer
    Dim x As Integer
    Dim vi
    x = 1
    For Each vi In lst
        If vi = v Then
            GetIndex = x
            Exit Function
        End If
        x = x + 1
    Next
End Function

Public Function IsTestOK(source As String, pattern As String) As Boolean
    Dim reg As New RegExp
    reg.pattern = pattern
    reg.IgnoreCase = True
    reg.MultiLine = True
    IsTestOK = reg.Test(source)
End Function

 
