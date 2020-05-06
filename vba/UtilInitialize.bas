Attribute VB_Name = "UtilInitialize"
Option Explicit

Public Sub Initialize()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
End Sub

Public Sub Finalize()
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub

'TODO ：ログ処理を追加する
