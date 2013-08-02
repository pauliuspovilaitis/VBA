Function ReturnBook(name As String) As Workbook
   Dim iSearch As Integer
   iSearch = 1
   Do While (arFailai(iSearch).name <> name) And (iSearch < FILE_ARRAY_SIZE - 1)
         iSearch = iSearch + 1
   Loop
   If (arFailai(iSearch).name = name) Then
        Set ReturnBook = arFailai(iSearch).book
   Else
        Set ReturnBook = Nothing
   End If

Function FindVerticalValue(book1 As String, Sheet1 As String, col As String, key As String) As Long
    Dim wBook1 As Workbook
    Set wBook1 = ReturnBook(book1)
    Dim iSearch As Long
    iSearch = 1
    Do While ((wBook1.Worksheets(Sheet1).Range(col & Val(iSearch)).Value <> key) And (iSearch < 65534))
       iSearch = iSearch + 1
    Loop
    If (iSearch = 65534) Then
        FindVerticalValue = 0
    Else
        FindVerticalValue = iSearch
    End If
  
    Set wBook1 = Nothing

End Function

Function FindHorizontalValue(book1 As String, Sheet1 As String, eilute As String, key As String) As String
    Dim wBook1 As Workbook
    Set wBook1 = ReturnBook(book1)
    Dim iSearch As Integer
    iSearch = 1
    Do While ((wBook1.Worksheets(Sheet1).Cells(eilute, Val(iSearch)).Value <> key) And (iSearch < 1000))
       iSearch = iSearch + 1
    Loop
    If (iSearch = 1000) Then
        FindHorizontalValue = ""
    Else
        FindHorizontalValue = ColumnLetter(iSearch)
    End If
  
    Set wBook1 = Nothing

End Function
