
Function FileOrDirExists(PathName As String) As Boolean
    
    Dim iTemp As Integer
     
    On Error Resume Next
    iTemp = GetAttr(PathName)
     
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select
     
    On Error GoTo 0
End Function

Sub main_checker()

    ThisWorkbook.Worksheets("main").Range("E13:E40").ClearContents
    
Dim dir As String
Dim filename As String
Dim dest_dir As String

Dim i As Integer

i = 13
Do While ThisWorkbook.Worksheets("MAIN").Range("B" & i) <> ""

    dir = ThisWorkbook.Worksheets("MAIN").Range("B" & i).Value
    filename = ThisWorkbook.Worksheets("MAIN").Range("C" & i).Value
    dest_dir = ThisWorkbook.Worksheets("MAIN").Range("D" & i).Value
    
    'check if found alreday
    If (ThisWorkbook.Worksheets("MAIN").Range("E" & i).Value <> "YES") Then
        'if no then label red
        ThisWorkbook.Worksheets("MAIN").Range("E" & i).Value = "NO :("
        ThisWorkbook.Worksheets("MAIN").Range("E" & i).Interior.Color = RGB(240, 15, 15)
        
        ' and check if exists already
        If (FileOrDirExists(dir & filename)) Then
        
            ThisWorkbook.Worksheets("MAIN").Range("E" & i).Value = "YES"
            ThisWorkbook.Worksheets("MAIN").Range("E" & i).Interior.Color = RGB(15, 240, 15)
    
    
            FileCopy dir & filename, dest_dir & filename
        End If
    End If
    

    i = i + 1
Loop


End Sub






