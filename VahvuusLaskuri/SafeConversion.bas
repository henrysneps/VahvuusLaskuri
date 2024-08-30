Attribute VB_Name = "SafeConversion"
Static Function ToInt(str As String) As Integer
    If IsNumeric(str) Then
        ToInt = CInt(str)
    Else
        ToInt = 0
    End If
End Function

Static Function ToString(intToConvert As Integer) As String
    ToString = CStr(intToConvert)
End Function
