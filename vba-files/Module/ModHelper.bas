Sub CreateSheetIfNotExists(sheetName As String)
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    
    ' �V�[�g�����݂��邩�m�F����
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' �V�[�g�����݂��Ȃ��ꍇ�A�V�����V�[�g���쐬����
    If Not sheetExists Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = sheetName
        Debug.Print "�V�����V�[�g '" & sheetName & "' ���쐬���܂����B"
    Else
        Debug.Print "�w�肵���V�[�g�� '" & sheetName & "' �͊��ɑ��݂��܂��B"
    End If
End Sub

' Sub TestCreateSheetIfNotExists()
'     CreateSheetIfNotExists "�V�����V�[�g"
' End Sub

Function GetTableFromCell(sheetName As String, cellAddress As String) As ListObject
    Dim tbl As ListObject
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        MsgBox "�w�肵���V�[�g�����݂��܂���B"
        Exit Function
    End If
    
    Set tbl = ws.Range(cellAddress).ListObject
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "�w�肵���Z���̓e�[�u���Ɋ܂܂�Ă��܂���B"
        Set GetTableFromCell = Nothing
    Else
        Set GetTableFromCell = tbl
    End If
End Function

' Sub TestGetTableFromCell()
'     Dim tbl As ListObject
'     Dim sheetName As String
'     Dim cellAddress As String
    
'     ' �V�[�g���ƃZ���̃A�h���X���w��
'     sheetName = "1"
'     cellAddress = "A1"
    
'     ' �e�[�u�����擾
'     Set tbl = GetTableFromCell(sheetName, cellAddress)
    
'     If Not tbl Is Nothing Then
'         MsgBox "�w�肵���Z�����܂܂��e�[�u���̖��O�� " & tbl.Name & " �ł��B"
'     End If
' End Sub

