Attribute VB_Name = "UtilSheetIn"
'*******************************************************
'UtilSheetIn
'Note:  �V�[�g���̑���
'*******************************************************
Option Explicit

'*****************************************
'����
'*****************************************
'*****************************************
'�擾
'*****************************************
'�󔒃Z���ɍi�荞��
Public Function FilterBlankCells(ByRef r As Range) As Range
On Error GoTo ErrHandl
    Set FilterBlankCells = r.SpecialCells(xlCellTypeBlanks)
    Exit Function
ErrHandl:
    '�󔒃Z�����������G���[����
    Set FilterBlankCells = Nothing
End Function

'�Z�����擾
Public Function GetSelectCellInfo(ByVal r As Range) As Object
On Error GoTo ErrHandl
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    
    Call dic.add("Column", r.Column)
    Call dic.add("Row", r.row)
    Call dic.add("Text", r.Text)
    Call dic.add("FColor", r.Font.color)
    Call dic.add("BColor", r.Interior.color)

    Set GetSelectCellInfo = dic
    Exit Function
ErrHandl:
    '�Z�����������G���[����
    Set GetSelectCellInfo = Nothing
End Function

'�V�F�C�v���擾
Public Function GetSelectShapeInfo(ByVal shp As Shape) As Object
On Error GoTo ErrHandl
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    
    Call dic.add("Text", shp.TextFrame.Characters.Text)
    Call dic.add("Pattern", shp.Fill.Pattern)
    Call dic.add("BColor", shp.Fill.BackColor)

    Set GetSelectShapeInfo = dic
    Exit Function
ErrHandl:
    '�Z�����������G���[����
    Set GetSelectShapeInfo = Nothing
End Function
'*****************************************
'�ݒ�
'*****************************************
'�S�ẴV�[�g�̃J�[�\����A1�Ɉړ�
Public Sub SetActiveA1()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim sht As Worksheet
    For Each sht In wb.Sheets
        sht.Activate
        sht.Range("A1").Select
    Next sht
    '�ŏ��Ɉړ�
    wb.Sheets(1).Activate
End Sub
'-----------------------------------------
'����
'-----------------------------------------
Public Sub SetShrinkToFit(ByRef r As Range, ByVal flg As Boolean)
    '�Z���̕�������Z�����ɍ��킹�Ă��ׂĕ\������
    r.ShrinkToFit = flg
End Sub

'-----------------------------------------
'�N���A
'-----------------------------------------
Public Sub SetClearOnlyFormat(ByRef r As Range)
    r.ClearFormats
End Sub
Public Sub SetClearOnlyVal(ByRef r As Range)
    r.ClearContents
End Sub
'-----------------------------------------
'�\��
'-----------------------------------------
'�W��
Public Sub SetNormalView()
    ActiveWindow.View = xlNormalView
End Sub
'���y�[�W�v���r���[
Public Sub SetPageBreakPreview()
    ActiveWindow.View = xlPageBreakPreview
End Sub
'���C�A�E�g�r���[
Public Sub SetPageLayoutView()
    ActiveWindow.View = xlPageLayoutView
End Sub
'-----------------------------------------
'�Y�[��
'-----------------------------------------
Public Sub SetZoom(ByVal zoomPer As Integer)
    ActiveWindow.Zoom = zoomPer
End Sub
'-----------------------------------------
'�O���[�v��/�A�E�g���C��
'-----------------------------------------
'�O���[�v�������ݒ�
Public Sub SetInitGroup(ByVal sht As Worksheet)
    With sht.Outline
        .AutomaticStyles = False
        '.SummaryRow = xlSummaryBelow        '�s�̃O���[�v���}�[�N��+/-��������(�W���ݒ�)
        '.SummaryColumn = xlSummaryOnRight   '��̃O���[�v���}�[�N��+/-���E����(�W���ݒ�)
        .SummaryRow = xlSummaryAbove        '�s�̃O���[�v���}�[�N��+/-���㑤��
        .SummaryColumn = xlSummaryOnLeft    '��̃O���[�v���}�[�N��+/-��������
    End With
End Sub

'�O���[�v���ݒ�
'41L~45L��I�����Đݒ聨40L�Ƀ}�[�N�o���B41L~45L�͂����܂��
Public Sub SetGroup(ByVal rng As Range)
    '�����ݒ�
    Call SetInitGroup(rng.Parent)
    rng.Rows.Group
End Sub

'�O���[�v������
Public Sub SetClearGroup(ByVal rng As Range)
'�O���[�v�������ꍇ�G���[���
On Error GoTo ErrHandl
    rng.Rows.Ungroup
    'rng.ClearOutline �������̂ق����ǂ��H
ErrHandl:
End Sub

'�A�E�g���C�������
Public Sub SetHideGroup(ByVal sht As Worksheet, ByVal colFlg As Boolean)
    If colFlg Then
        sht.Outline.ShowLevels ColumnLevels:=1
    Else
        sht.Outline.ShowLevels RowLevels:=1
    End If
End Sub

'�A�E�g���C�����J��
Public Sub SetOpenGroup(ByVal sht As Worksheet, ByVal colFlg As Boolean)
    If colFlg Then
        sht.Outline.ShowLevels ColumnLevels:=2
    Else
        sht.Outline.ShowLevels RowLevels:=2
    End If
End Sub

