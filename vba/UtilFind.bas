Attribute VB_Name = "UtilFind"
Option Explicit
'*****************************************
'�擾
'*****************************************
'�������ʃZ����Address���X�g��Ԃ�
Public Function SetFind(ByVal rng As Range, ByVal str As String, ByVal findFormarFlg As Boolean) As Collection
    '�ŏ�������
    Dim searchRes As Range
    Set searchRes = rng.Find(What:=str, _
                            SearchFormat:=findFormarFlg, _
                            SearchDirection:=xlNext)
    '�Ȃ���ΏI��
    If searchRes Is Nothing Then
        Set SetFind = Nothing
        Exit Function
    End If
    '�ŏ��̏ꏊ���L��
    Dim firstResAdd As String: firstResAdd = searchRes.Address
    
    '�ŏ��̌������ʂɂȂ�܂Ń��[�v
    Dim res As New Collection
    Do
        res.add searchRes.Address
        '��������
        Set searchRes = rng.FindNext(searchRes)
    Loop Until searchRes.Address = firstResAdd
    
    Set SetFind = res
End Function
'*****************************************
'�ݒ�
'*****************************************
'-----------------------------------------
'��������
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
