VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private ExportModules As Boolean

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim ActionToTake&

ActionToTake = MsgBox("Do you want to export VBA Code upon successful save?", vbYesNoCancel, Me.Name)

Select Case ActionToTake
  Case vbYes
    ExportModules = True
  Case vbNo
    ExportModules = False
  Case Else
    Cancel = True
End Select

End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)

If Success Then
  If ExportModules Then
    Call ExportVisualBasicCode
  End If
End If

End Sub

Private Function SameText(ByVal String1 As String, ByVal String2 As String, ByVal CaseSensitive As Boolean) As Boolean

String1 = Trim(String1)
String2 = Trim(String2)

If Not CaseSensitive Then
  String1 = StrConv(String1, vbUpperCase)
  String2 = StrConv(String2, vbUpperCase)
End If

SameText = (StrComp(String1, String2, vbTextCompare) = 0)

End Function
Private Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim Path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    'Begin Mods for error described below
    Dim FolderPicker As FileDialog
    
    'Below line causes error due to OneDrive folders are in a web url.
    'directory = ActiveWorkbook.path & "\VisualBasic"
        
    'To fix this, use a Dialog to pick the folder
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    GetLocalPath (Application.ActiveWorkbook.Path) & Application.PathSeparator
    With FolderPicker
      .InitialFileName = GetLocalPath(Application.ActiveWorkbook.Path) & Application.PathSeparator
      .Title = "Select a Folder"
      .AllowMultiSelect = False
      If .Show <> -1 Then Exit Sub
      directory = .SelectedItems(1) & "\Source"
    End With
    'End mods
    
    count = 0
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        Path = directory & Application.PathSeparator & VBComponent.Name & extension
        Call VBComponent.Export(Path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & Path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & Path
        End If

        On Error GoTo 0
    Next
    
    'Not needed and causes annoying warning. Replace with MsgBox Output.
    'Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    'Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
    MsgBox ("Successfully exported " & CStr(count) & " VBA files to " & directory)
    
End Sub

Private Function GetLocalPath(ByVal FullPath As String) As String
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


