Attribute VB_Name = "UtilDriver"
'*******************************************************
'UtilDriver
'Note:  �o�^�p�}�N��
'*******************************************************

'@Folder("VBAProject")
Option Explicit

'�t�b�^�[�ݒ�
Public Sub test()
    Call SetPrintFooter(ActiveSheet, "", "&P/&N�y�[�W", "&F")
End Sub

'�y�[�W�ݒ�
Public Sub test2()

    With ActiveWorkbook
        .Worksheets.Select
        Call SetPageBreakPreview
        Call SetZoom(80)
        .Worksheets(1).Select
    End With
End Sub

'�I��͈͂̋󔒃Z�����������Z�b�g
Public Sub SetClearFormatBlankCells()
    Initialize
    Dim r As Range: Set r = FilterBlankCells(Selection)
    Call SetClearOnlyFormat(r)
    Finalize
End Sub

'-----------------------------------------
'���
'-----------------------------------------
Public Sub �J�[�\��A1�ړ�()
    Call SetActiveA1
End Sub
'-----------------------------------------
'����
'-----------------------------------------
Public Sub �k���\��OFF()
    ' �Z���̕�������Z�����ɍ��킹�Ă��ׂĕ\������(���{���ł͐ݒ�ł��Ȃ�)
    Call SetShrinkToFit(ActiveCell, False)
End Sub
Public Sub �k���\��ON()
    Call SetShrinkToFit(ActiveCell, True)
End Sub

'-----------------------------------------
'�F
'-----------------------------------------
Public Sub ������()
    Call SetFontBlack
End Sub

Public Sub ������()
    Call SetFontRed
End Sub

Public Sub ������()
    Call SetFontBlue
End Sub

Public Sub �w�iC()
    Call SetBackClear
End Sub

Public Sub �w�i��()
    Call SetBackRed
End Sub

Public Sub �w�i��()
    Call SetBackBlue
End Sub

Public Sub �w�i��()
    Call SetBackYellow
End Sub

Public Sub �w�i�D()
    Call SetBackGray
End Sub

'-----------------------------------------
'�f�[�^�擾
'-----------------------------------------
Public Sub GetSelectCell()
    Dim dic As Object
    Set dic = GetSelectCellInfo(ActiveCell)
    If dic Is Nothing Then
        MsgBox "�Z�����I������Ă��܂���"
        Exit Sub
    End If
    
    Dim key As Variant
    For Each key In dic.Keys
        Debug.Print "Key: " & key, "Value: " & dic.Item(key)
    Next
End Sub

Public Sub GetSelectShape()
    Dim shp As Shape
    For Each shp In Selection.ShapeRange
        
        '�擾
        Dim dic As Object
        Set dic = GetSelectShapeInfo(shp)
        If dic Is Nothing Then
            MsgBox "�}�`���I������Ă��܂���"
            Exit Sub
        End If
        
        '�\��
        Dim key As Variant
        For Each key In dic.Keys
            Debug.Print "Key: " & key, "Value: " & dic.Item(key)
        Next
        
        dic = Nothing
    Next
End Sub

'-----------------------------------------
'�O���[�v���ŃA�E�g���C��
'-----------------------------------------
Public Sub SetGroupByMethod()
    Dim sht As Worksheet: Set sht = ActiveSheet
    '������
    Call SetClearGroup(sht.UsedRange)
    
    '----------------------
    '�O���[�v���̋�؂������
    '----------------------
    '���������ݒ�
    Call SetFindFormatClear
    Call SetFindFormatBackColor(CLR_SKYBLUE)
    'P��ɔw�i�F���F�́u���\�b�h���v������
    Dim colMethodAddList As Collection
    Set colMethodAddList = SetFind("P3:P" & GetEndRow(sht), "���\�b�h��", True)
    '��������������
    Call SetFindFormatClear
    '�����ꍇ�͏I��
    If colMethodAddList Is Nothing Then
        Exit Sub
    End If
    
    '----------------------
    '�O���[�v�ݒ�
    '----------------------
    Dim i As Integer
    For i = 1 To colMethodAddList.Count
        '���\�b�h��`�s 1�s��
        Dim selCellAdd As String
        selCellAdd = colMethodAddList.Item(i)
        '�w�i�F���F�łȂ��Ȃ�ꏊ�܂ŉ����
        Dim tmpR As Range: Set r = sht.Range(selCellAdd)
        Do
            '����̍s�ɉ����
            Set tmpR = tmpR.Offset(1, 0)
        Loop Until tmpR.Interior.color <> CLR_SKYBLUE
        
        '�O���[�v���J�n�s
        Dim grpSttCell As Range: Set grpSttCell = tmpR
        '�O���[�v���ŏI�s
        Dim grpEndCell As Range
        If i = colMethodAddList.Count Then
            '�Ō�̒�`�̏ꍇ
            Set grpEndCell = sht.Range("P" & GetEndRow(sht))
        Else
            '���̒�`�̈�s�O
            Set grpEndCell = sht.Range(colMethodAddList.Item(i + 1)).Offset(-1, 0)
        End If
        
        '�O���[�v��
        Call SetGroup(sht.Range(grpSttCell, grpEndCell))
    Next
End Sub

Public Sub ��G�S��()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim sht As Worksheet
    For Each sht In wb.Sheets
        Call SetHideGroup(sht, True)
    Next
End Sub

Public Sub ��G�S�J()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim sht As Worksheet
    For Each sht In wb.Sheets
        Call SetOpenGroup(sht, True)
    Next
End Sub

