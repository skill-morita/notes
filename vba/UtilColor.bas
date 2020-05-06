Attribute VB_Name = "UtilColor"
Option Explicit

'*****************************************
'判定
'*****************************************
Public Function IsRed(ByVal colorVal As Long) As Boolean
    If colorVal = CLR_RED Then
        IsRed = True
        Exit Function
    End If
    IsRed = False
End Function

'*****************************************
'取得
'*****************************************
'-----------------------------------------
'RGB
'-----------------------------------------
Public Function GetRedByRGB(ByVal colorVal As Long) As Long
    GetRedByRGB = colorVal Mod 256
End Function

Public Function GetGreenByRGB(ByVal colorVal As Long) As Long
    GetGreenByRGB = Int(colorVal / 256) Mod 256
End Function

Public Function GetBlueByRGB(ByVal colorVal As Long) As Long
    GetBlueByRGB = Int(colorVal / 256 / 256)
End Function

'*****************************************
'設定
'*****************************************
'-----------------------------------------
'フォント色
'-----------------------------------------
Public Sub SetFontBlack()
    Application.Selection.Font.color = CLR_BLACK
End Sub

Public Sub SetFontRed()
    Application.Selection.Font.color = CLR_RED
End Sub

Public Sub SetFontBlue()
    Application.Selection.Font.color = CLR_SCUMBLUE
End Sub

'-----------------------------------------
'背景色
'-----------------------------------------
Public Sub SetBackClear()
    Application.Selection.Interior.colorIndex = CLR_IDX_CLEAR
End Sub

Public Sub SetBackRed()
    Application.Selection.Interior.color = CLR_RED
End Sub

Public Sub SetBackBlue()
    Application.Selection.Interior.color = CLR_SKYBLUE
End Sub

Public Sub SetBackYellow()
    Application.Selection.Interior.color = CLR_YELLOW
End Sub

Public Sub SetBackGray()
    Application.Selection.Interior.color = CLR_LIGHTGRAY
End Sub
