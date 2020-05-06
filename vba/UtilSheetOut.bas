Attribute VB_Name = "UtilSheetOut"
'*******************************************************
'UtilSheetOut
'Note:  シート外の操作
'*******************************************************
Option Explicit

'*****************************************
'判定
'*****************************************
'*****************************************
'取得
'*****************************************
'タブ色が指定色のシートリスト取得
Public Function GetColorSheet(ByVal colorVal As Long) As Collection
    Dim res As New Collection
    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Sheets
        If sht.Tab.color = colorVal Then
            res.add sht
        End If
    Next
    
    Set GetColorSheet = res
End Function
'*****************************************
'設定
'*****************************************


