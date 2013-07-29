'additional functions
Public Function IsWorkBookOpen(WkbName As String) As Boolean
    Dim wBook As Workbook

    On Error Resume Next
    Set wBook = Workbooks(WkbName)
    
    If (wBook Is Nothing) Then
        Set wBook = Nothing
        IsWorkBookOpen = False
    On Error GoTo 0
    Else
        Set wBook = Nothing
        IsWorkBookOpen = True
    On Error GoTo 0
    End If
End Function

Public Function GetLong(gText As String) As Long
    GetLong = Len(gText)
    Do While (Mid(gText, GetLong, 1) <> "\")
        GetLong = GetLong - 1
    Loop
End Function
