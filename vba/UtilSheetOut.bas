Attribute VB_Name = "UtilSheetOut"
'*******************************************************
'UtilSheetOut
'Note:  �V�[�g�O�̑���
'*******************************************************
Option Explicit

'*****************************************
'����
'*****************************************
'*****************************************
'�擾
'*****************************************
'�^�u�F���w��F�̃V�[�g���X�g�擾
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
'�ݒ�
'*****************************************


