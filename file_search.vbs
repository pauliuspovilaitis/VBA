Sub find_xls_files(mysourcepath)

    per_row = 2
    
    Set myobject = New Scripting.filesystemobject
    Set mysource = myobject.GetFolder(mysourcepath)
    
        For Each myfile In mysource.Files
            If Right(myfile, 5) = ".xlsx" Or Right(myfile, 4) = ".xls" Then
                col = 2
                ThisWorkbook.Worksheets("found_files").Cells(per_row, col).Value = myfile.Path
                ThisWorkbook.Worksheets("found_files").Cells(per_row, col + 1).Value = myfile.Name
            End If
            per_row = per_row + 1
        Next
        
End Sub
