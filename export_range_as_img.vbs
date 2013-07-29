Sub export_as_image(export_dir As String, sh As String, rg as String)

ThisWorkbook.Worksheets(sh).Activate

With ThisWorkbook.Worksheets(sh)
    
    On Error Resume Next
    
    Dim rgExp As Range: Set rgExp = Range(rg)
    rgExp.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    

    With ActiveSheet.ChartObjects.Add(Left:=rgExp.Left, Top:=rgExp.Top, _
        Width:=rgExp.Width, Height:=rgExp.Height)
        .Name = "main"
        .Activate
    End With
    
    ActiveChart.Paste
    ActiveSheet.ChartObjects(sh).Chart.Export export_dir
    ActiveSheet.ChartObjects(sh).Delete
    
 End With
  
End Sub
