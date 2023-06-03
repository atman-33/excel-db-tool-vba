VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public TabName As String
Public TableComment As String

Public ColumnNames As New Collection
Public DataRows As New Collection
Public ErrorMessage As String

Public Sub Fill(source As String)
    RemoveAll ColumnNames
    RemoveAll DataRows
    
    If Left(source, 6) = "error:" Then
        ErrorMessage = source
        Exit Sub
    Else
        ErrorMessage = ""
    End If
    
    Dim rows() As String
    rows = Split(source, vbCrLf)
    Dim row
    Dim DataRow As Collection
    Dim i As Integer
    For Each row In rows
        If row <> "" Then
            If i = 0 Then
                'header
                Call SplitDataToCollection(CStr(row), ColumnNames, True)
            Else
                'data
                Set DataRow = New Collection
                Call SplitDataToCollection(CStr(row), DataRow, False)
                DataRows.Add DataRow
            End If
        End If
        '=====================
        i = i + 1
    Next
End Sub
