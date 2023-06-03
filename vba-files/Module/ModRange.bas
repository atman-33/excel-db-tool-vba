Attribute VB_Name = "ModRange"
Option Explicit


'アプリ名を取得
Public Function GetAppName(appid As String, findrange As Range, Optional col As Integer = 1) As String
    Dim thisRange As Range
    Set thisRange = GetRangeValue(appid, findrange)
    If Not thisRange Is Nothing Then
        GetAppName = thisRange.Worksheet.Cells(thisRange.row, thisRange.Column + col).Text
    Else
        GetAppName = ""
    End If
End Function

'指定値を指定範囲に検索
Public Function GetRangeValue(findValue, findrange As Range, Optional row = 0, Optional col = 0) As Range
    Dim thisRange As Range
    Set thisRange = findrange.Find(findValue)
    If Not thisRange Is Nothing Then
        Set GetRangeValue = thisRange.Cells(row + 1, col + 1)
    End If
End Function

'指定値を指定範囲に検索
Public Function findrange(thisRange As Range, findKey As String, Optional afterCell As Range = Nothing) As Range
    Dim thisSHeet As Worksheet
    Set thisSHeet = thisRange.Worksheet
    If afterCell Is Nothing Then
        Set afterCell = thisSHeet.Cells(thisRange.row, thisRange.Column)
    End If
    
    Set findrange = thisRange.Find(What:=findKey, After:=afterCell, lookAt:=xlPart)
End Function

'指定値を指定シートに検索
Public Function FindInSheet(thisSHeet As Worksheet, findKey As String, Optional afterCell As Range = Nothing) As Range
    Dim thisRange As Range
    If afterCell Is Nothing Then
        Set afterCell = thisSHeet.Cells(1, 1)
    End If
    
    Set FindInSheet = thisSHeet.Cells.Find(What:=findKey, After:=afterCell, lookAt:=xlPart)
End Function


Public Sub InsertRow(thisSHeet As Worksheet, r As Integer, count As Integer)
    thisSHeet.rows(r & ":" & (r + count)).Insert Shift:=xlDown
    thisSHeet.rows(r & ":" & (r + count)).Clear
End Sub

Public Sub DeleteRow(thisSHeet As Worksheet, r1 As Integer, r2 As Integer)
    If r1 >= r2 Then Exit Sub
    
    thisSHeet.rows(r1 & ":" & r2).Delete Shift:=xlUp
    thisSHeet.Cells(r1, 1).Select
End Sub
Public Sub ShowHeaderRangeStyle(thisSHeet As Worksheet, r1 As Integer, c1 As Integer, r2 As Integer, c2 As Integer)
    ShowRangeStyle thisSHeet, r1, c1, r2, c2, 20
End Sub

Public Sub ShowDataRangeStyle(thisSHeet As Worksheet, r1 As Integer, c1 As Integer, r2 As Integer, c2 As Integer)
    ShowRangeStyle thisSHeet, r1, c1, r2, c2, 2
End Sub

Public Sub ShowRangeStyle(thisSHeet As Worksheet, r1 As Integer, c1 As Integer, r2 As Integer, c2 As Integer, colorindex As Integer)
    Dim headRange As Range
    Dim cell1 As Range
    Dim cell2 As Range
    Dim maxCol As Integer
    Dim newMaxCol As Integer
   
    Set cell1 = thisSHeet.Cells(r1, c1)
    Set cell2 = thisSHeet.Cells(r2, c2)
    Set headRange = thisSHeet.Range(cell1, cell2)
    SetLineStyle headRange, xlContinuous
    SetLineWeight headRange, xlThin
    headRange.Interior.colorindex = colorindex
End Sub

'指定範囲に□線の太さを設定
Public Sub SetLineWeight(theRange As Range, weight As XlBorderWeight)
    theRange.Borders(xlEdgeTop).weight = weight
    theRange.Borders(xlEdgeBottom).weight = weight
    theRange.Borders(xlEdgeLeft).weight = weight
    theRange.Borders(xlEdgeRight).weight = weight
    theRange.Borders(xlInsideVertical).weight = weight
    theRange.Borders(xlInsideHorizontal).weight = weight
End Sub
'指定範囲に□線を設定
Public Sub SetLineStyle(theRange As Range, lineStyle As XlLineStyle)
    theRange.Borders(xlEdgeTop).lineStyle = lineStyle
    theRange.Borders(xlEdgeBottom).lineStyle = lineStyle
    theRange.Borders(xlEdgeLeft).lineStyle = lineStyle
    theRange.Borders(xlEdgeRight).lineStyle = lineStyle
    theRange.Borders(xlInsideVertical).lineStyle = lineStyle
    theRange.Borders(xlInsideHorizontal).lineStyle = lineStyle
End Sub
Sub ShowColorIndex(a)

    Dim r As Integer
    r = 1
    For r = 1 To 30
        ActiveSheet.Cells(r, 1).Interior.colorindex = r
    Next

End Sub
