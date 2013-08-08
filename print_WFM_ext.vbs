Sub EXPORT()

  Dim MyFile As Variant
  MyFile = ThisWorkbook.Worksheets("Settings").Range("b9").Value
    
    fnum = FreeFile()
    Open MyFile For Output As fnum
       
    Dim eilute As String
    
    Dim printer As Long
    printer = 2
    
    
    With ThisWorkbook.Worksheets("data")
    
        Do While ThisWorkbook.Worksheets("data").Range("a" & printer).Value <> ""
            
            eilute = Format(.Cells(printer, 1).Value, "mm/dd/yy") & Chr(32) & _
            Format(.Cells(printer, 2).Value, "hh:mm") & Chr(32) & _
            .Cells(printer, 3).Value & Chr(32) & _
            .Cells(printer, 5).Value & Chr(32) & _
            .Cells(printer, 6).Value & Chr(32) & _
            .Cells(printer, 7).Value & Chr(32) & _
            Right("00000000" & Format(.Cells(printer, 8).Value, "0.00"), 11) & Chr(32) & _
            Right("00000000" & Format(.Cells(printer, 9).Value, "0.00"), 11) & Chr(32) & _
            Right("00000000" & Format(.Cells(printer, 10).Value, "0.0"), 7) & Chr(32) & _
            Right("00000000" & Format(.Cells(printer, 11).Value, "0.00"), 8) & Chr(32) & _
            .Cells(printer, 12).Value
            
            Print #fnum, eilute
               
        printer = printer + 1
        Loop
        
   End With
   
   Print #fnum, "$END OF ASPECT"
   
  Close #fnum
    
    
End Sub
