Attribute VB_Name = "ModPlugin"
Public Function ReplaceKeyValue(origin As String, key As String, value As String) As String

    ReplaceKeyValue = Application.Run("ReplaceKeyValue", origin, key, value)

End Function

Public Function ReplaceFiles(c1 As String, c2 As String, c3 As String, c4 As String) As String
    
    Call Application.Run("ReplaceFiles", c1, c2, c3, 0, c4)

End Function
