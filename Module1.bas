Public Sub mcrImportVBACode()

Dim FilesPicker As FileDialog
Dim Filename As Variant
Dim FileType As Variant
    
Set FilesPicker = Application.FileDialog(msoFileDialogOpen)

With FilesPicker
  .Title = "Select Files to Import"
  .InitialFileName = GetLocalPath(Application.ActiveWorkbook.Path) & Application.PathSeparator
  .Filters.Add "VBA Source Files", "*.cls; *.frm; *.bas", 1
  .AllowMultiSelect = True
End With
    
If FilesPicker.Show Then
  For Each Filename In FilesPicker.SelectedItems
    ActiveWorkbook.VBProject.VBComponents.Import Filename
    DoEvents
  Next
End If

Set FilesPicker = Nothing

End Sub
