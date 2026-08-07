Option Explicit

Private originalCalculation As XlCalculation

'/*
'Optimizes execution speed by disabling UI screen rendering,
'events, and automatic calculation.
'*/
Public Sub EnableOptimization()
    originalCalculation = Application.Calculation
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    On Error GoTo 0
End Sub

'/*
'Restores UI rendering, events, alerts, and original
'calculation state
'*/
Public Sub DisableOptimization()
    On Error Resume Next
    If originalCalculation = 0 Then originalCalculation = xlCalculationAutomatic
    Application.Calculation = originalCalculation
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub