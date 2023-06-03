VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsMetaTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public TableName As String
Public TableComment As String

Private metaData As New ClsDataTable
Public DataColumns As New Collection
Public ParamMap As New Dictionary


Public OrderBy As String
Public RowNum As Integer

Public Property Get ColumnNames() As Collection
    Set ColumnNames = metaData.ColumnNames
End Property

Public Sub Init(TableName As String, dbinfo As ClsDBInfo)
    
    Me.TableName = TableName
    
    ParamMap.RemoveAll
    ParamMap.Add "TNAME", TableName
    dbinfo.AddMap ParamMap
    
    Dim strSQL As String
    strSQL = GetSQL("GetTableComment", ParamMap)
    If strSQL = "" Then
        MsgBox "SQL取得失敗、処理終了。", vbExclamation
        End
    End If
    'テーブル和名
    Me.TableComment = CStr(dbinfo.GetObject(strSQL))
            
    strSQL = GetSQL("GetTableColumnsInfo", ParamMap)
    If strSQL = "" Then
        MsgBox "SQL取得失敗、処理終了。", vbExclamation
        End
    End If
    
    '列情報
    Dim strResult As String
    strResult = dbinfo.Query(strSQL)
    
    If IsTestOK(strResult, "error") Then
        MsgBox strResult, vbExclamation
        End
    End If
    metaData.Fill strResult
    
    RemoveAll DataColumns
    '列情報を初期化
    Dim row As Collection
    Dim ci As ClsColumnInfo
    For Each row In metaData.DataRows
        Set ci = New ClsColumnInfo
        ci.CName = row(ColumnIndex("CNAME"))
        ci.ColType = row(ColumnIndex("COLTYPE"))
        ci.ScaleW = row(ColumnIndex("SCALE"))
        ci.Width = row(ColumnIndex("WIDTH"))
        ci.Nulls = row(ColumnIndex("NULLS"))
        ci.Pkey = row(ColumnIndex("CONSTRAINT_TYPE"))
        ci.comment = row(ColumnIndex("COMMENTS"))
        ci.DefaultValue = row(ColumnIndex("DEFAULTVAL"))
        If IsTestOK(ci.comment, "^$") Then
            ci.comment = ci.CName
        End If
        DataColumns.Add ci
    Next
    
End Sub

Public Function ColumnIndex(colName As String) As Integer
    ColumnIndex = GetIndex(metaData.ColumnNames, colName)
End Function

Public Function GetSelectSQL(thisSHeet As Worksheet, r As Integer, Optional dblinkName As String = "") As String
    
    Dim ci As ClsColumnInfo
    For Each ci In Me.DataColumns
        '該当列の現在値設定
        ci.SetCValue thisSHeet, r
    Next
    
    Dim fromTable As String
    If dblinkName <> "" Then
        fromTable = Me.TableName + "@" + dblinkName
    Else
        fromTable = Me.TableName
    End If
    
    GetSelectSQL = " SELECT " + GetColumnsString _
                + " FROM " + fromTable _
                + " WHERE " + GetWhereString _
                + GetOrderByString
End Function

Public Function GetDeleteSQL(thisSHeet As Worksheet, r As Integer, Optional dblinkName As String = "") As String
    'データ設定
    SetColumnValue thisSHeet, r
    
    Dim fromTable As String
    If dblinkName <> "" Then
        fromTable = Me.TableName + "@" + dblinkName
    Else
        fromTable = Me.TableName
    End If
    
    
    GetDeleteSQL = " DELETE FROM " + fromTable _
                + " WHERE " + GetPKeyWhereString
End Function

Public Function GetInsertSQL(thisSHeet As Worksheet, r As Integer, Optional dblinkName As String = "") As String
    'データ設定
    SetColumnValue thisSHeet, r
    
    Dim fromTable As String
    If dblinkName <> "" Then
        fromTable = Me.TableName + "@" + dblinkName
    Else
        fromTable = Me.TableName
    End If
    
    
    GetInsertSQL = " INSERT INTO " + fromTable + "(" + GetColumnsString + ")" _
                + " VALUES(" + GetInsertValueString + ")"
End Function

Private Function GetColumnsString() As String
    Dim selecCols As String
    Dim ci As ClsColumnInfo
    For Each ci In Me.DataColumns
        '選択対象列
        selecCols = selecCols + "," + ci.CName
    Next
    GetColumnsString = Mid(selecCols, 2)
End Function

Private Function GetWhereString() As String
    Dim where As String
    Dim ci As ClsColumnInfo
    For Each ci In Me.DataColumns
        '検索条件
        If ci.CValue <> "" Then
            If IsTestOK(ci.CValue, "NULL") _
            Or IsTestOK(ci.CValue, "\(") _
            Or IsTestOK(ci.CValue, "\s+AND") _
            Or IsTestOK(ci.CValue, "LIKE\s+") Then
                where = where + " AND " + ci.CName + " " + ci.CValue
            Else
                where = where + " AND " + ci.CName + "=" + ci.CValue
            End If
            
        End If
    Next
     '検索条件
    If CStr(Me.RowNum) <> "" Then
        where = where + " AND ROWNUM<=" + CStr(Me.RowNum)
    End If
    
    GetWhereString = Mid(where, 5)
End Function

Private Function GetPKeyWhereString() As String
    Dim where As String
    Dim ci As ClsColumnInfo
    For Each ci In Me.DataColumns
        '検索条件
        If ci.Pkey = "P" Then
            where = where + " AND " + ci.CName + " = " + ci.CValue
        End If
    Next
    
    GetPKeyWhereString = Mid(where, 5)
End Function

Private Function GetInsertValueString() As String
    Dim insertValue As String
    Dim ci As ClsColumnInfo
    For Each ci In Me.DataColumns
        '新規対象
        If ci.CValue <> "" Then
            insertValue = insertValue + "," + ci.CValue
        ElseIf ci.Nulls = "NULL" Then
            insertValue = insertValue + ",NULL"
        Else
            insertValue = insertValue + "," + ci.DefaultValue
        End If
    Next
    
    GetInsertValueString = Mid(insertValue, 2)
End Function

Private Function GetUpdateString() As String
    Dim updateColumns As String
    Dim ci As ClsColumnInfo
    For Each ci In Me.DataColumns
        '更新対象
        If ci.CValue <> "" Then
            updateColumns = updateColumns + "," + ci.CName + " = " + ci.CValue
        ElseIf ci.Nulls = "NULL" Then
            updateColumns = updateColumns + "," + ci.CName + " = NULL"
        Else
            updateColumns = updateColumns + "," + ci.CName + " = " + ci.DefaultValue
        End If
    Next
    GetUpdateString = Mid(updateColumns, 2)
End Function
Private Function GetOrderByString() As String
    Dim orderByString As String
    '並び順
    If Me.OrderBy <> "" Then
        orderByString = " ORDER BY " + Me.OrderBy
    End If
    
    GetOrderByString = orderByString
End Function

Public Sub SetColumnValue(thisSHeet As Worksheet, r As Integer)
    Dim ci As ClsColumnInfo
    For Each ci In Me.DataColumns
        '該当列の現在値設定
        ci.SetCValue thisSHeet, r
    Next
End Sub
