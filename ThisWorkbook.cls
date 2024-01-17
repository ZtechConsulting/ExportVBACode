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
