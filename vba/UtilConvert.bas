Attribute VB_Name = "UtilConvert"
Option Explicit

'Collection -> 1�����z��
Public Function ColToAry(ByVal targetCol As Collection) As Variant
 
    Dim resAry() As Variant
    '�z��̃T�C�Y�ݒ�
    If targetCol.Count <> 0 Then
        ReDim resAry(targetCol.Count - 1)
    End If
     
    'Collection�̊e�v�f��z��ɃZ�b�g
    Dim index As Integer: index = 0
    Dim itm As Variant
    For Each itm In targetCol
        resAry(index) = itm
        index = index + 1
    Next
  
    '�Ԃ�
    ColToAry = resAry
End Function

'1�����z�� -> Collection
Public Function AryToCol(ByVal ary As Variant) As Collection
    Dim res As New Collection
    
    Dim itm As Variant
    For Each itm In ary
        Call res.add(itm)
    Next
    Set AryToCol = res
End Function


'1�����z�� -> 2�����z��

