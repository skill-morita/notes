Attribute VB_Name = "UtilConvert"
Option Explicit

'Collection -> 1次元配列
Public Function ColToAry(ByVal targetCol As Collection) As Variant
 
    Dim resAry() As Variant
    '配列のサイズ設定
    If targetCol.Count <> 0 Then
        ReDim resAry(targetCol.Count - 1)
    End If
     
    'Collectionの各要素を配列にセット
    Dim index As Integer: index = 0
    Dim itm As Variant
    For Each itm In targetCol
        resAry(index) = itm
        index = index + 1
    Next
  
    '返す
    ColToAry = resAry
End Function

'1次元配列 -> Collection
Public Function AryToCol(ByVal ary As Variant) As Collection
    Dim res As New Collection
    
    Dim itm As Variant
    For Each itm In ary
        Call res.add(itm)
    Next
    Set AryToCol = res
End Function


'1次元配列 -> 2次元配列

