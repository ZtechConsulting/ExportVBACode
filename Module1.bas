Attribute VB_Name = "Module1"
Option Explicit

Public Function GetLocalPath(ByVal FullPath As String) As String
    'Finds local path for a OneDrive file URL, using environment variables of OneDrive
    'Reference https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive
    'Authors: Philip Swannell 2019-01-14, MatChrupczalski 2019-05-19, Horoman 2020-03-29, P.G.Schild 2020-04-02
    Dim ii&
    Dim iPos&
    Dim OneDrivePath$
    Dim endFilePath$
    Dim NbSlash&
    
    'If Left(Trim(FullPath, 8)) = "https://" Then
    If SameText(Left(Trim(FullPath), 8), "https://", False) Then
        If InStr(1, FullPath, "sharepoint.com/") <> 0 Then 'Commercial OneDrive
            NbSlash = 4
        Else 'Personal OneDrive
            NbSlash = 2
        End If
        iPos = 8 'Last slash in https://
        For ii = 1 To NbSlash
            iPos = InStr(iPos + 1, FullPath, "/")
        Next ii
        endFilePath = Mid(FullPath, iPos)
        endFilePath = Replace(endFilePath, "/", Application.PathSeparator)
        For ii = 1 To 3
            OneDrivePath = Environ(Choose(ii, "OneDriveCommercial", "OneDriveConsumer", "OneDrive"))
            If 0 < Len(OneDrivePath) Then Exit For
        Next ii
        GetLocalPath = OneDrivePath & endFilePath
        While Len(Dir(GetLocalPath, vbDirectory)) = 0 And InStr(2, endFilePath, Application.PathSeparator) > 0
            endFilePath = Mid(endFilePath, InStr(2, endFilePath, Application.PathSeparator))
            GetLocalPath = OneDrivePath & endFilePath
        Wend
    Else
        GetLocalPath = FullPath
    End If
End Function

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

Public Function SameText(ByVal String1 As String, ByVal String2 As String, ByVal CaseSensitive As Boolean) As Boolean

String1 = Trim(String1)
String2 = Trim(String2)

If Not CaseSensitive Then
  String1 = StrConv(String1, vbUpperCase)
  String2 = StrConv(String2, vbUpperCase)
End If

SameText = (StrComp(String1, String2, vbTextCompare) = 0)

End Function
