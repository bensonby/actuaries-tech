Private VersionString As String = "0.1"

'Required inputs
Public InputListHeading As Range
Public CellToPasteInput As Range
Public ResultCellsToCopy As Range
Public ColumnHeadingToPasteResultCells As Range 'TODO: how about dynamic range? does it work?

'Optional inputs
Public IsRestoreInputCell As Boolean = True
Public IsRestoreSystemSettings As Boolean = False

Public StartingRowNumber As Long = 1
Public EndingRowNumber As Long = 0 '0 means auto-determined
Public PivotTableNamesToRefresh() As String = Array()
Public IsAppendResults As Boolean = True 'otherwise overwrite mode
Public IsPasteFormat As Boolean = False 'otherwise require system clipboard

Public RunStatusColumnHeading As Range 'show completed time??
Public ExtraFilterColumnHeading As Range '
Public ExtraFilterValueToRun As String = "Y"
Public IsIgnoreErrors As Boolean = False '"True" is not recommended - all output cells should be without errors. Errors indicate problems

Public IsShowProgress As Boolean = True
Public IsClearResults As Boolean = False '"False" is not recommended - may lead to undesired user experience

Public IsScreenUpdating As Boolean = False 'if false, keep current config
Public IsManualCalc As Boolean = False '"False" is not recommended - effective spreadsheet design for data table should behave the same with auto calculation; if false, keep current config

'System variables
Private BackupInputCellValue As String

Public Function Run()
  Dim IsValid As Boolean
  IsValid = Me.CheckAndInitialiseInputs()
  If Not IsValid Then Exit Sub

  Dim IsConfirm As Boolean
  IsConfirm = Me.PromptInputs() 'show my URL of the library
  If Not IsConfirm Then Exit Sub

  If IsRestoreInputCell Then Call Me.BackupInputCell
  If IsRestoreSystemSettings Then Call Me.BackupSystemSettings

  Call Me.SetSystemSettings
  If IsClearResults Then Call Me.ClearResultsBeforeRun
  Call Me.RunDataTable
  Call Me.ShowCompleteDialog

  If Me.IsRestoreInputCell Then Call Me.RestoreInputCell
  If Me.IsRestoreSystemSettings Then Call Me.RestoreSystemSettings
End Sub

Private Sub ClearResultsBeforeRun()
End Sub

Private Sub SetSystemSettings()
  If IsScreenUpdating Then Application.ScreenUpdating = False
  If IsManualCalc Then Application.Calculation = xlCalculationManual
End Sub

Private Sub BackupSystemSettings()
  'TODO
End Sub

Private Sub BackupInputCell()
  BackupInputCellValue = CellToPasteInput.Value
End Sub

Private Function PromptInputs() As Boolean
  Dim message As String = Me.get_summary_string()
  toContinue = MsgBox( message & "... OK?", vbYesNo, "Continue?")
  If toContinue = vbNo Then PromptInputs = False
  Else PromptInputs = True
End Function

Private Function GetSummaryString() As String
  GetSummaryString = "Running records #" & CStr(StartingRowNumber) & " to #" & CStr(EndingRowNumber) & VbCrLf & _
                      & "blah blah blah" & VbCrLf
  'TODO: get other config to display!
End Function

Private Function CheckAndInitialiseInputs() As Boolean
  Dim ErrorMessages(0 to 10) As String
  
  If InputListHeading is Nothing OR CellToPasteInput is Nothing OR _
     ResultCellsToCopy is Nothing OR ColumnHeadingToPasteResultCells is Nothing Then
  End If

  If StartingRowNumber < 1 Then
  End If

  If EndingRowNumber <> 0 And StartingRowNumber > EndingRowNumber Then
  End If

  If StartingRowNumber is ""...
  End If

  If PivotTableNamesToRefresh Is Not Nothing Then
    For i = LBound(PivotTableNamesToRefresh) To UBound(PivotTableNamesToRefresh)
      If Not IsValidPivotTableName(PivotTableNamesToRefresh(i) Then 'TODO
      End If
    Next i
  End If

  'If No Errors
  If something Then
    Call SetEndingRowNumber
  End If

End Function
