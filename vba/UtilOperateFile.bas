Attribute VB_Name = "UtilOperateFile"
Option Explicit

'�t�H���_���̃t�@�C���擾
Public Function GetFiles(ByVal FolderPath As String) As Collection
    '�t�@�C���V�X�e���I�u�W�F�N�g�擾
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Folder�I�u�W�F�N�g�擾
    Dim f As Object
    Set f = fso.GetFolder(FolderPath)

    'File�R���N�V������Ԃ�
    Dim res As New Collection
    Dim itm As Variant
    For Each itm In f.Files
        res.add itm.Path
    Next
    
    Set GetFiles = res
End Function

'�w��Z���̓��e��CSV�ɏo�͂���
'�o�̓p�X���擾
'�I���Z�����e���擾
'selection.cells(1.1).text
'�t�@�C���o��


'*********************************************
'Excel
'*********************************************
Public Function OpenExcel(ByVal filepath As String, ByVal readOnlyFlg As Boolean) As Workbook
    Dim res As Workbook
    Set res = Workbooks.Open(filepath, readOnlyFlg)
End Function

Public Sub CloseExcel(ByRef wb As Workbook, ByVal saveChangeFlg As Boolean)
    Call wb.Close(saveChangeFlg)
End Sub

'�l�}�N���u�b�N�ȊO�̊J���Ă���u�b�N���擾����
Public Function GetOpenBook() As Collection
    Dim res As New Collection
    Dim wb As Workbook
    For Each wb In Workbooks
        If MACRO_BOOK <> wb.name Then
            res.add wb
        End If
    Next
    Set GetOpenBook = res
End Function

