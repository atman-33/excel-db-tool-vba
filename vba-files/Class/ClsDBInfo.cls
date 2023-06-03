VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDBInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Host As String
Public Port As String
Public SID As String
Public Owner As String
Public Password As String
Public DBLink As String

Const OracleConnectString = "User ID={OWNER}; Password={PASSWORD}; Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=tcp)(HOST={HOST})(PORT={PORT}))(CONNECT_DATA=(SERVICE_NAME={SID})))"

Private Property Get ConnectString()
    
    ConnectString = OracleConnectString
    ConnectString = Replace(ConnectString, "{OWNER}", Owner)
    ConnectString = Replace(ConnectString, "{PASSWORD}", Owner)
    ConnectString = Replace(ConnectString, "{HOST}", Host)
    ConnectString = Replace(ConnectString, "{PORT}", Port)
    ConnectString = Replace(ConnectString, "{SID}", SID)
End Property

Public Sub AddMap(dbMap As Dictionary)
    dbMap.Add "HOST", Host
    dbMap.Add "PORT", Port
    dbMap.Add "SID", SID
    dbMap.Add "OWNER", Owner
    dbMap.Add "PASSWORD", Password
    dbMap.Add "DBLINK", DBLink
End Sub

Public Sub Init(thisSHeet As Worksheet)
    If thisSHeet.Range("A1").Text <> "HOST" Then
        MsgBox "DBÚ‘±—pHOSTî•ñ‚ðŠi”[‚µ‚Ä‚¢‚é³‹K‚ÈƒV[ƒg‚Å‚Í—L‚è‚Ü‚¹‚ñ.", vbExclamation, "ˆ—I—¹"
        End
    End If
    
    Host = thisSHeet.Range("B1")
    Port = thisSHeet.Range("D1")
    SID = thisSHeet.Range("F1")
    Owner = thisSHeet.Range("H1")
    Password = thisSHeet.Range("J1")
    DBLink = thisSHeet.Range("L1")
End Sub


Public Function Batch(strSQL As String) As String
    Dim sql() As String
    sql = Split(strSQL, ";")
    Dim s
    For Each s In sql
        ToLog CStr(s)
    Next
    
    Batch = Application.Run("Batch", strSQL, ConnectString)
    If Err.Number > 0 Then
        ToLog ConnectString & "; " & Err.Description
    End If
End Function

Public Function Query(strSQL As String) As String
    ToLog strSQL
    
    Query = Application.Run("Query", strSQL, ConnectString)
    
    If Err.Number > 0 Then
        ToLog ConnectString & "; " & Err.Description
    End If
End Function
Public Function GetObject(strSQL As String) As Variant
    ToLog strSQL
    
    GetObject = Application.Run("GetObject", strSQL, ConnectString)
    
    If Err.Number > 0 Then
        ToLog ConnectString & "; " & Err.Description
    End If
End Function


