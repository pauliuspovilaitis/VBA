On Error Resume Next
  Dim fso, folder, files, NewsFile, sFolder, einamasis
  
  Const ForReading = 1
  Const ForWriting = 2
 
  Set fso = CreateObject("Scripting.FileSystemObject")
  sFolder = "D:\Scripts\NetFlix\Agent_Activity\macro_output"
  If sFolder = "" Then
      Wscript.Echo "No Folder parameter was passed"
      Wscript.Quit
  End If

  Set folder = fso.GetFolder(sFolder)
  Set files = folder.Files
 
  For each folderIdx In files
	If UCase(fso.GetExtensionName(folderIdx.Name)) = "CSV" Then
		Set einamasis = fso.OpenTextFile(folderIdx.path, ForReading) 
		AllTheText = einamasis.ReadAll
		AllTheText = Replace(AllTheText, """", "")
		AllTheText = Replace(AllTheText, chr(13), "")
		AllTheText = Replace(AllTheText, chr(10) & chr(10), chr(10))
		einamasis.Close     
		Set einamasis = fso.OpenTextFile(folderIdx.path, ForWriting)
		If (Right(AllTheText,2) = Chr(10) & Chr(10)) Then
			AllTheText = Replace(AllTheText, chr(10) & chr(10), chr(10))
		End If
		einamasis.Write AllTheText
		einamasis.Close
	End If
  Next
