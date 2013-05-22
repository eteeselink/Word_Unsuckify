Attribute VB_Name = "Unsuckify"
Sub RelativeImage()

    docPath = ActiveDocument.Path

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.InitialFileName = docPath
    ok = fd.Show
    If ok = 0 Then
      Exit Sub
    End If
    
    fullPath = fd.SelectedItems(1)
    If InStr(fullPath, docPath) <> 1 Then
      MsgBox ("Image must be in the same directory or a subdirectory of the document")
      Exit Sub
    End If
    
    ' We do Len(docPath) + 2 because we need 1 for VB's off by one shit,
    ' and 1 for the trailing backslash
    relativePath = Mid(fullPath, Len(docPath) + 2)
    relativePath = Replace(relativePath, "\", "\\")
    
    ' Add fields as if they're typed, because that is the only simple way to
    ' make nested fields (we nest the FILENAME field inside the INCLUDEPICTURE field)
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False
    Selection.TypeText ("INCLUDEPICTURE """)
    
    ' Make nested FILENAME field
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False
    Selection.TypeText ("FILENAME \p")
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
    Selection.TypeText ("\\..\\" & relativePath & """ \* MERGEFORMAT")
    Selection.Fields.Update
End Sub
