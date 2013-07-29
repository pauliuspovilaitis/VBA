Sub NewZip(sPath)
'empty zip file creation
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub


Sub zip(FolderName As Variant, DefPath As String, name As String)
Dim FileNameZip
    Dim oApp As Object
FileNameZip = DefPath & name & ".zip"
  
    NewZip (FileNameZip)

    Set oApp = CreateObject("Shell.Application")
    'Copy the files to zip
    oApp.Namespace(FileNameZip).CopyHere oApp.Namespace(FolderName).items

    'keep the script wait while working...
    On Error Resume Next
    Do Until oApp.Namespace(FileNameZip).items.count = _
       oApp.Namespace(FolderName).items.count
        Application.Wait (Now + TimeValue("0:00:15"))
    Loop
    On Error GoTo 0


End Sub
