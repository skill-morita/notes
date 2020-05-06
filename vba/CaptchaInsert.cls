VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaptchaInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Type ImgInfo
    filepath As String
    Per As Integer
        
    left As Single
    Top As Single
    
    Height As Single
    Width As Single
End Type

Public Sub CaptchaInsert()
    Const FolderPath As String = "C:\temp\"
    Const Ext As String = "png"
    Dim interval As Single
    interval = 5

    ' �w��t�H���_�z���̃t�@�C�����X�g�擾
    Dim pathList As Collection
    Set pathList = GetFiles(FolderPath)
    
    ' ���݈ʒu�擾
    Dim posL As Single
    Dim posT As Single
    posL = Selection.left
    posT = Selection.Top
    
    For Each p In pathList
        ' �g���q���t�@�C���p�X�Ɋ܂܂�Ă���
        If InStr(p.Path, Ext) > 0 Then
            Dim iInfo As ImgInfo
            iInfo.filepath = p.Path
            iInfo.left = posL
            iInfo.Top = posT
            iInfo.Per = 100
            
            ' �摜�t�@�C���̑}��
            iInfo = InsertImg(ActiveSheet, iInfo)
            ' �摜�̃T�C�Y�����Ɉړ�����
            posT = posT + iInfo.Top + interval
            
            ' �}�������摜�͕s�v�Ȃ̂ō폜
            Kill p.Path
        End If
    Next
 
End Sub


'�摜��}������
Private Function InsertImg(ByRef sht As Worksheet, ByRef iInfo As ImgInfo) As ImgInfo
    Dim shp As Shape
    
    ' �摜�\��t��
    With sht
        Set shp = .Shapes.AddPicture( _
                Filename:=iInfo.filepath, _
                linktofile:=False, _
                savewithdocument:=True, _
                left:=iInfo.left, _
                Top:=iInfo.Top, _
                Width:=0, _
                Height:=0)
    End With

    With shp
        ' ��U�c���̃T�C�Y��100%�Ƀ��Z�b�g����
        .LockAspectRatio = False
        .ScaleHeight 1, msoTrue, msoScaleFromTopLeft
        .ScaleWidth 1, msoTrue, msoScaleFromTopLeft

        ' �c����Œ�Ŏw�聓�ɐݒ�
        .LockAspectRatio = True
        .ScaleHeight (iInfo.Per / 100), msoTrue, msoScaleFromTopLeft
        
        ' �őO�ʂɔz�u
        .ZOrder msoBringToFront
        
        ' �摜�T�C�Y�擾
        iInfo.Height = .Height
        iInfo.Width = .Width
    End With

    Set shp = Nothing
    
    ' �摜�T�C�Y�Ԃ�
    InsertImg = iInfo
End Function


