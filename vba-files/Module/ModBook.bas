Attribute VB_Name = "ModBook"
Option Explicit
Public fso As New FileSystemObject

'================================================
' Workbook SaveAs
'================================================
Function BookSaveAs(thisBook As Workbook, strFileName As String) As Integer
On Error GoTo err1:

    thisBook.SaveAs strFileName, CreateBackup:=False
    Exit Function
    
err1:
    BookSaveAs = -1
End Function
'================================================
' Open Workbook
'================================================
Function OpenBook(strFileName As String) As Workbook
    Dim book As Workbook
    
    For Each book In Workbooks
        If book.Name = strFileName Then
            Set OpenBook = book
            Exit Function
        End If
        If book.FullName = strFileName Then
            Set OpenBook = book
            Exit Function
        End If
    Next
    
    If fso.FileExists(strFileName) Then
        Application.DisplayAlerts = False
        Set book = Workbooks.Open(strFileName, readonly:=False, IgnoreReadOnlyRecommended:=True, Editable:=True)
    Else
        Set book = Workbooks.Add()
        BookSaveAs book, strFileName
    End If
    If book.Sheets.count > 1 And HasSheet("Sheet1", book) Then
        book.Sheets("Sheet1").Delete
    End If
    Set OpenBook = book
End Function
'================================================
' Workbook SaveAs
'================================================
Sub BookCloseAndSave(thisBook As Workbook, Optional fullpath As String = "")
Dim showAlert As Boolean
    
    showAlert = Application.DisplayAlerts
    
    Application.DisplayAlerts = False
    On Error Resume Next
    If fullpath = "" Then
        thisBook.Close True
    Else
        BookSaveAs thisBook, fullpath
        thisBook.Close
    End If
    Application.DisplayAlerts = showAlert
    
End Sub

'================================================
' Workbook Close by Name
'================================================
Sub BookCloseAndSaveByName(workbookName As String)
    Dim book As Workbook
    For Each book In Workbooks
        If book.Name = workbookName Then
            BookCloseAndSave book
            Exit Sub
        End If
        If book.FullName = workbookName Then
            BookCloseAndSave book
            Exit Sub
        End If
    Next
    
End Sub

'================================================
' Workbook Close by Names
'================================================
Sub BookCloseAndSaveByNames(workbookNames() As String)
    Dim i%
    For i = 0 To UBound(workbookNames)
       BookCloseAndSaveByName workbookNames(i)
    Next
    
End Sub

'================================================
' Close Workbook Without Save and no alert.
'================================================
Function BookCloseNoSave(book As Workbook) As Boolean
    
    Dim showAlert As Boolean
    
    showAlert = Application.DisplayAlerts
    
    Application.DisplayAlerts = False
    book.Close False
    Application.DisplayAlerts = showAlert
    
End Function

Public Function HasSheet(sheetName As String, Optional book As Workbook = Nothing) As Boolean
    If book Is Nothing Then
        Set book = ThisWorkbook
    End If
    
    Dim sheet As Worksheet
    For Each sheet In book.Sheets
        If LCase(sheet.Name) = LCase(sheetName) Then
            HasSheet = True
            Exit Function
        End If
    Next
End Function

Public Function GetSheet(sheetName As String, Optional book As Workbook = Nothing) As Worksheet
    If book Is Nothing Then
        Set book = ThisWorkbook
    End If
    
    If HasSheet(sheetName, book) Then
        Set GetSheet = book.Sheets(sheetName)
        Exit Function
    End If
    
    Dim sheet As Worksheet
    Set sheet = book.Sheets.Add
    sheet.Name = sheetName
    Set GetSheet = sheet
End Function
 
Public Sub CopyData(fromSheet As Worksheet, toSheet As Worksheet, r As Integer, Optional title As String = "")
    Dim rx, cx As Integer
    rx = fromSheet.UsedRange.rows.count
    cx = fromSheet.UsedRange.Columns.count
    Dim fromRange As Range
    Set fromRange = fromSheet.Range(fromSheet.Cells(1, 1), fromSheet.Cells(rx, cx))
    fromRange.Copy
    
    toSheet.Activate
    toSheet.Range("B" & r).Select
    toSheet.Paste
    toSheet.Range("A" & r) = title
    If 1 < rx Then
        toSheet.Range("A" & r & ":A" & (r + rx - 1)).FillDown
    End If
End Sub

Public Sub ShowSheetTable(thisSHeet As Worksheet)
    
    Dim r, c As Integer
    r = thisSHeet.UsedRange.rows.count
    c = thisSHeet.UsedRange.Columns.count
    Dim thisRange As Range
    Set thisRange = thisSHeet.Range(thisSHeet.Cells(1, 1), thisSHeet.Cells(r, c))
    
    thisSHeet.ListObjects.Add(xlSrcRange, thisRange, , xlYes).Name = thisSHeet.Name
    thisSHeet.ListObjects(thisSHeet.Name).TableStyle = "TableStyleLight21"
    
End Sub


Public Sub ExportReport(thisSHeet As Worksheet, reportFolder As String)
    'Save To XlsReport
    If Not fso.FolderExists(reportFolder) Then
        fso.CreateFolder reportFolder
    End If
    
    SaveAsNewXls thisSHeet, fso.BuildPath(reportFolder, thisSHeet.Name & Format(Now, "YYYYMMDDHHMMSS") & ".xlsx")
    
End Sub

Sub SaveAsNewXls(thisSHeet As Worksheet, reportPath As String)
    
    thisSHeet.Activate
    thisSHeet.Copy
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs reportPath
    ActiveWindow.Close
    Application.DisplayAlerts = True
End Sub


'Public Function findrange(thisRange As Range, findKey As String, Optional afterCell As Range = Nothing, Optional lookAt As Integer = xlPart) As Range
'    Dim thisSHeet As Worksheet
'    Set thisSHeet = thisRange.Worksheet
'    If afterCell Is Nothing Then
'        Set afterCell = thisSHeet.Cells(thisRange.row, thisRange.Column)
'    End If
'
'    Set findrange = thisRange.Find(What:=findKey, After:=afterCell, lookAt:=lookAt)
'End Function

Public Function FindInSheet(thisSHeet As Worksheet, findKey As String, Optional afterCell As Range = Nothing) As Range
    Dim thisRange As Range
    If afterCell Is Nothing Then
        Set afterCell = thisSHeet.Cells(1, 1)
    End If
    
    Set FindInSheet = thisSHeet.Cells.Find(What:=findKey, After:=afterCell, lookAt:=xlPart)
End Function

Public Function GetFilePath(para As String) As String
    Dim filePath$
    filePath = ThisWorkbook.Path
    If (para <> "") Then
        filePath = filePath + "\" + para
    End If
    GetFilePath = filePath
End Function
