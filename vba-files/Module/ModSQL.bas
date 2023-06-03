Attribute VB_Name = "ModSQL"
Const COMMENT_TAB = "COMMENT ON TABLE {TNAME} IS '{COMMENT}';"
Const COMMENT_COL = "COMMENT ON COLUMN {TNAME}.{CNAME} IS '{COMMENT}'; "

Public SQLList As New Dictionary

Private Sub RefreshSQLList()
    'èâä˙âª
    SQLList.RemoveAll
    Dim r As Integer
    Dim c As Integer
    r = 2
    c = 3
    Dim o  As ClsSQL
    '
    Do While Not IsEmpty(Sheet10.Cells(r, c))
        Set o = New ClsSQL
        o.id = Sheet10.Cells(r, c).Text
        o.sql = Sheet10.Cells(r, c + 1).Text
        o.parameters = Sheet10.Cells(r, c + 2).Text
        o.ReturnType = Sheet10.Cells(r, c + 3).Text
        SQLList.Add o.id, o
        '==================
        r = r + 1
        DoEvents
    Loop
End Sub

Public Function GetSQL(id As String, params As Dictionary) As String
    If SQLList.count = 0 Then
        RefreshSQLList
    End If
    
    If SQLList.Exists(id) Then
        Dim o As ClsSQL
        Set o = SQLList(id)
        Dim strSQL As String
        strSQL = o.sql
        If o.parameters <> "" Then
            Dim p
            For Each p In params
                strSQL = ReplaceKeyValue(strSQL, CStr(p), params(p))
            Next
        End If
        GetSQL = strSQL
    Else
        GetSQL = ""
    End If
End Function

Public Sub ToLog(msg1 As String, Optional msg2 As String = "", Optional msg3 As String = "")
    Dim msg As String
    msg = Format(Now, "yyyy/MM/dd hh:mm:ss")
    msg = msg & vbTab & msg1
    msg = msg & vbTab & msg2
    msg = msg & vbTab & msg3
    
    Dim logPath As String
    logPath = fso.BuildPath(ThisWorkbook.Path, Replace(ThisWorkbook.Name, ".xlsm", ".log"))
    Dim ts As TextStream
    If fso.FileExists(logPath) Then
        Set ts = fso.OpenTextFile(logPath, ForAppending)
    Else
        Set ts = fso.CreateTextFile(logPath, ForAppending)
    End If
    ts.WriteLine msg
    ts.Close
    
End Sub
'Sub toCommentsql()
'
'    Dim thisbook As Workbook
'    Set thisbook = ActiveWorkbook
'
'    Dim comment As String
'    Dim thisSHeet As Worksheet
'    Dim re As New RegExp
'    re.pattern = "(\w+)\s*\((.+)\)"
'    Dim ms As MatchCollection
'
'    For Each thisSHeet In thisbook.Worksheets
'        Set ms = re.Execute(thisSHeet.Range("A4"))
'        If ms.count > 0 Then
'            comment = COMMENT_TAB
'            comment = Replace(comment, "{TNAME}", ms(0).SubMatches(0))
'            comment = Replace(comment, "{COMMENT}", ms(0).SubMatches(1))
'
'            toColsql thisSHeet, ms(0).SubMatches(0)
'        End If
'        DoEvents
'    Next
'
'End Sub
'
'Private Sub toColsql(thisSHeet As Worksheet, TableName As String)
'
'Dim comment As String
'Dim r As Integer
'r = 7
'While Not IsEmpty(thisSHeet.Cells(r, 1))
'    comment = COMMENT_COL
'    comment = Replace(comment, "{TNAME}", TableName)
'    comment = Replace(comment, "{CNAME}", thisSHeet.Range("C" & r))
'    comment = Replace(comment, "{COMMENT}", thisSHeet.Range("B" & r))
'
'    r = r + 1
'    DoEvents
'Wend
'End Sub
'
