VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Const ROWS_HEADER_FROM = 7
Const COLS_HEADER_FROM = 2
Const ROWS_HEADER_COUNT = 5

Public MetaTable As New ClsMetaTable
Public DataTable As New ClsDataTable

Public Function Showed(thisSHeet As Worksheet) As Boolean
    Showed = (thisSHeet.Range("B2") = thisSHeet.Range("L2")) _
            And (thisSHeet.Range("B2") = MetaTable.TableName) _
            And MetaTable.ColumnNames.count > 0
End Function

Public Sub ShowMetaTable(thisSHeet As Worksheet)
    On Error GoTo err1
    
    If Showed(thisSHeet) Then
        Exit Sub
    End If
    'エラーメッセージ領域をクリア
    thisSHeet.Range("B3").ClearContents
    'DB接続情報取得
    dbinfo.Init thisSHeet
    
    '基本情報取得
    MetaTable.Init thisSHeet.Range("B2"), dbinfo
    
    'テーブル名表示
    thisSHeet.Range("D2") = MetaTable.TableComment
    'テーブル列情報表示
    ShowHeaderStyle thisSHeet
    ShowHeadColumn thisSHeet
    
    'テーブル情報表示済
    thisSHeet.Range("L2") = thisSHeet.Range("B2")
    Exit Sub
err1:
    MsgBox Err.Description
End Sub

Public Sub ShowDataRows(thisSHeet As Worksheet)
    
    On Error GoTo err1
    'データエリア初期化
    ClearDataRows thisSHeet
    
    '検索条件、並び順設定
    MetaTable.TableName = thisSHeet.Range("B2")
    MetaTable.RowNum = thisSHeet.Range("F2")
    MetaTable.OrderBy = thisSHeet.Range("H2")
    
    DataTable.ErrorMessage = ""
    '検索処理
    Dim strSQL As String
    strSQL = MetaTable.GetSelectSQL(thisSHeet, ROWS_HEADER_FROM, dbinfo.DBLink)
    thisSHeet.Range("C3") = strSQL
    
    Dim strResult As String
    strResult = dbinfo.Batch(strSQL)
    
    DataTable.Fill strResult
    If DataTable.ErrorMessage <> "" Then
        thisSHeet.Range("B3") = DataTable.ErrorMessage
        thisSHeet.Range("B3").Interior.colorindex = 3
        ToLog thisSHeet.Name, DataTable.ErrorMessage
    Else
        thisSHeet.Range("B3") = ""
        thisSHeet.Range("B3").Interior.colorindex = 2
        ToLog thisSHeet.Name, "SQL Select OK."
    End If
    
    Dim row As Collection
    Dim ci As ClsColumnInfo
    Dim r As Integer
    r = ROWS_HEADER_FROM + ROWS_HEADER_COUNT + 1
    
    For Each row In DataTable.DataRows
        For Each ci In MetaTable.DataColumns
            thisSHeet.Cells(r, ci.Index) = row(ci.Index - COLS_HEADER_FROM + 1)
        Next
        '行番号累計
        r = r + 1
    Next
    
    Exit Sub
err1:
    MsgBox Err.Description
End Sub


Private Sub ShowHeaderStyle(thisSHeet As Worksheet)
    Dim headRange As Range
    Dim cell1 As Range
    Dim cell2 As Range
    Dim maxCol As Integer
    Dim newMaxCol As Integer
    maxCol = thisSHeet.UsedRange.Columns.count
    If maxCol < MetaTable.DataColumns.count + COLS_HEADER_FROM - 1 Then
        maxCol = MetaTable.DataColumns.count + COLS_HEADER_FROM - 1
    End If
    
    Set cell1 = thisSHeet.Cells(ROWS_HEADER_FROM + 1, COLS_HEADER_FROM)
    Set cell2 = thisSHeet.Cells(ROWS_HEADER_FROM + ROWS_HEADER_COUNT, maxCol)
    Set headRange = thisSHeet.Range(cell1, cell2)
    headRange.Clear
    newMaxCol = MetaTable.DataColumns.count + COLS_HEADER_FROM - 1
    Set cell2 = thisSHeet.Cells(ROWS_HEADER_FROM + ROWS_HEADER_COUNT, newMaxCol)
    Set headRange = thisSHeet.Range(cell1, cell2)
    SetLineStyle headRange, xlContinuous
    SetLineWeight headRange, xlThin
    headRange.Interior.colorindex = 20
End Sub


Private Sub ClearDataRows(thisSHeet As Worksheet)
    Dim headRange As Range
    Dim cell1 As Range
    Dim cell2 As Range
    Dim maxRow As Integer
    Dim maxCol As Integer
    maxRow = thisSHeet.UsedRange.rows.count
    If maxRow < ROWS_HEADER_FROM + ROWS_HEADER_COUNT + 1 Then
        maxRow = ROWS_HEADER_FROM + ROWS_HEADER_COUNT + 1
    End If
    maxCol = thisSHeet.UsedRange.Columns.count
    If maxCol < MetaTable.DataColumns.count + COLS_HEADER_FROM - 1 Then
        maxCol = MetaTable.DataColumns.count + COLS_HEADER_FROM - 1
    End If
    
    Set cell1 = thisSHeet.Cells(ROWS_HEADER_FROM + ROWS_HEADER_COUNT + 1, 1)
    Set cell2 = thisSHeet.Cells(maxRow, maxCol)
    Set headRange = thisSHeet.Range(cell1, cell2)
    headRange.ClearContents
End Sub

Private Sub ShowHeadColumn(thisSHeet As Worksheet)
    
    Dim ci As ClsColumnInfo
    Dim c As Integer
    c = COLS_HEADER_FROM
    Dim r As Integer
    r = ROWS_HEADER_FROM + 1
    For Each ci In MetaTable.DataColumns
        '========
        'WHERE行は検索条件として、保留
        '========
        'PKEY,'コメント
        If ci.Pkey = "P" Then
            '背景色設定
            thisSHeet.Cells(r, c) = ci.comment & "(PK)"
            thisSHeet.Cells(r, c).Interior.colorindex = 4
        Else
            thisSHeet.Cells(r, c) = ci.comment
        End If
        
        '列英名
        thisSHeet.Cells(r + 1, c) = ci.CName
        '型名
        thisSHeet.Cells(r + 2, c) = ci.ColType
        '桁数
        If ci.ScaleW = "" Then
            thisSHeet.Cells(r + 3, c) = ci.Width
        Else
            thisSHeet.Cells(r + 3, c) = ci.Width & "," & ci.ScaleW
        End If
        'NULLS
        thisSHeet.Cells(r + 4, c) = ci.Nulls
        '所在列
        ci.Index = c
    '=============================
        c = c + 1
    Next
    
End Sub

Public Sub UpdateData(thisSHeet As Worksheet)
    
    Dim ci As ClsColumnInfo
    Dim r As Integer
    r = ROWS_HEADER_COUNT + ROWS_HEADER_FROM + 1
    
    Dim Result As String
    Dim strSQL As String
    While Not IsEmpty(thisSHeet.Range("B" & r))
        '削除
        If strSQL <> "" Then
            strSQL = strSQL + ";"
        End If
        strSQL = strSQL + MetaTable.GetDeleteSQL(thisSHeet, r, dbinfo.DBLink)
        
        '新規登録
        If IsEmpty(thisSHeet.Range("A" & r)) Then
            strSQL = strSQL & ";" & MetaTable.GetInsertSQL(thisSHeet, r, dbinfo.DBLink)
        End If
        '===========================
        r = r + 1
        DoEvents
    Wend
    thisSHeet.Range("C3") = Replace(strSQL, vbLf, " ")
    Result = dbinfo.Batch(strSQL)
    
    If IsTestOK(Result, "error") Then
        thisSHeet.Range("B3") = Result
        thisSHeet.Range("B3").Interior.colorindex = 3
        ToLog thisSHeet.Name, Result
        Exit Sub
    End If
    ToLog thisSHeet.Name, "Update Data OK."
    
End Sub
