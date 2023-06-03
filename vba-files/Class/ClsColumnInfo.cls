VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsColumnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public CName As String
Public ColType As String
Public ScaleW As String
Public Width As String
Public Nulls As String
Public Pkey As String
Public comment As String
Public DefaultValue As String

Public Index As Integer

'åªç›íl
Public CValue As String


Public Sub SetCValue(thisSHeet As Worksheet, r As Integer)
    
    Dim strDate As String
    Dim ci As ClsColumnInfo
    Dim cellText As String
    Set ci = Me
    
   
    cellText = thisSHeet.Cells(r, ci.Index).Text
    ci.CValue = ""
    If cellText = "" Then
        Exit Sub
    End If
    
    'åüçıèåè
    If ci.Nulls = "NULL" And IsTestOK(cellText, "IS\s*(NOT)*\s*NULL") Then
        ci.CValue = cellText
    ElseIf IsTestOK(cellText, "IN\s*\([\s\w,']+\)") _
    Or IsTestOK(cellText, "BETWEEN\s+[\w']+\s+AND\s+[\w']+") _
    Or IsTestOK(cellText, "LIKE\s+") Then
        'IN('a', 'b', 'c'), BETWEEN 1 AND 10
        ci.CValue = cellText
    ElseIf cellText = "(null)" Then
        ci.CValue = "null"
    Else
        Select Case ci.ColType
        Case "NUMBER"
            ci.CValue = cellText
        Case "DATE"
            If IsTestOK(cellText, "\d{4}\/\d{2}\/\d{2}\s+\d+:\d+:\d+") Then
                strDate = "TO_DATE('" + cellText + "','YYYY/MM/DD HH24:MI:SS')"
            ElseIf IsTestOK(cellText, "\d{4}\/\d{2}\/\d{2}") Then
                strDate = "TO_DATE('" + cellText + "','YYYY/MM/DD')"
            ElseIf IsTestOK(cellText, "\d{4}\d{2}\d{2}") Then
                strDate = "TO_DATE('" + cellText + "','YYYYMMDD')"
            ElseIf IsTestOK(cellText, "\d{4}\d{2}\d{2}\d+\d+\d+") Then
                strDate = "TO_DATE('" + cellText + "','YYYYMMDDHH24MISS')"
            Else
                strDate = "TO_DATE('" + cellText + "')"
            End If
            
            ci.CValue = strDate
            
        Case "TIMESTAMP"
            If IsTestOK(cellText, "\d{4}\/\d{2}\/\d{2}\s+\d+:\d+:\d+.\d+") Then
                strDate = "TO_DATE('" + cellText + "','YYYY/MM/DD HH24:MI:SS.FFF')"
            ElseIf IsTestOK(cellText, "\d{4}\d{2}\d{2}\d{2}\d+\d+\d+") Then
                strDate = "TO_DATE('" + cellText + "','YYYYMMDDHH24MISSFFF')"
            Else
                strDate = "TO_TIMESTAMP('" + cellText + "')"
            End If
            
            ci.CValue = strDate
        
        Case Else '             "CHAR", "VARCHAR", "VARCHAR2"
            ci.CValue = "'" + cellText + "'"
        End Select
    End If

End Sub


