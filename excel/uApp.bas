Option Explicit

Private originalCalculation As XlCalculation

'/*
' Disables screen updating, events, alerts, and automatic calculation.
'*/
Public Sub EnableOptimization()
    On Error Resume Next
    With Application
        originalCalculation = .Calculation
        .Calculation = xlCalculationManual
        .DisplayStatusBar = False
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    ActiveSheet.DisplayPageBreaks = False
    On Error GoTo 0
End Sub

'/*
' Restores screen updating, events, alerts, and calculation state.
'*/
Public Sub DisableOptimization(Optional ByVal restoreOriginalCalc As Boolean = True)
    On Error Resume Next
    With Application
        If Not restoreOriginalCalc Or originalCalculation = 0 Then
            .Calculation = xlCalculationAutomatic
        Else
            .Calculation = originalCalculation
        End If
        .DisplayStatusBar = True
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
    On Error GoTo 0
End Sub