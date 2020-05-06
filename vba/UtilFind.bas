Attribute VB_Name = "UtilFind"
Option Explicit
'*****************************************
'取得
'*****************************************
'検索結果セルのAddressリストを返す
Public Function SetFind(ByVal rng As Range, ByVal str As String, ByVal findFormarFlg As Boolean) As Collection
    '最初を検索
    Dim searchRes As Range
    Set searchRes = rng.Find(What:=str, _
                            SearchFormat:=findFormarFlg, _
                            SearchDirection:=xlNext)
    'なければ終了
    If searchRes Is Nothing Then
        Set SetFind = Nothing
        Exit Function
    End If
    '最初の場所を記憶
    Dim firstResAdd As String: firstResAdd = searchRes.Address
    
    '最初の検索結果になるまでループ
    Dim res As New Collection
    Do
        res.add searchRes.Address
        '次を検索
        Set searchRes = rng.FindNext(searchRes)
    Loop Until searchRes.Address = firstResAdd
    
    Set SetFind = res
End Function
'*****************************************
'設定
'*****************************************
'-----------------------------------------
'検索書式
'-----------------------------------------
Public Sub SetFindFormatClear()
    Application.FindFormat.Clear
End Sub

Public Sub SetFindFormatFontColor(ByVal colorVal As Long)
    Application.FindFormat.Font.color = colorVal
End Sub

Public Sub SetFindFormatBackColor(ByVal colorVal As Long)
    Application.FindFormat.Interior.color = colorVal
End Sub
