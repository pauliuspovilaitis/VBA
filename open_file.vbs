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

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Dim file3 As String
    file3 = ThisWorkbook.Worksheets("settings").Range("c6").Value
    
    Dim file3_obj As Workbook
    Dim file3_name As String
    file3_name = Right(file3, Len(file3) - GetLong(file3))
    If (IsWorkBookOpen(file3_name) = False) Then
        Set file3_obj = Workbooks.Open(Filename:=file3)
    Else
        Set file3_obj = Workbooks(file3_name)
    End If
