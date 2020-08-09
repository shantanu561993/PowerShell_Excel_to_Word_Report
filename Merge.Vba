Sub Merge()
  Dim dlgFile As FileDialog
  Dim nTotalFiles As Integer
  Dim nEachSelectedFile As Integer

  Set dlgFile = Application.FileDialog(msoFileDialogFilePicker)
 
  With dlgFile
    .AllowMultiSelect = True
    If .Show <> -1 Then
      Exit Sub
    Else
      nTotalFiles = .SelectedItems.Count
    End If
  End With
 
  For nEachSelectedFile = 1 To nTotalFiles
    Selection.InsertFile dlgFile.SelectedItems.Item(nEachSelectedFile)
    If nEachSelectedFile < nTotalFiles Then
      Selection.InsertBreak Type:=wdPageBreak
    Else
      If nEachSelectedFile = nTotalFiles Then
        Exit Sub
      End If
    End If
  Next nEachSelectedFile
End Sub
