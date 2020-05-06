Attribute VB_Name = "CountPage"
Option Explicit

'TODO : �V�F�C�v�̐ԕ����擾

'*****************************************
'1
'*****************************************
'�u�b�N�S�̂̃y�[�W���擾
Public Sub CheckBookExistRed()
    Initialize

    Dim col As New Collection
    Dim allCnt As Long
    
    '�V�[�g�^�u���Ԃ̂��̂��擾
    Dim shtCol As Collection: Set shtCol = GetColorSheet(CLR_RED)
    
    Dim sht As Worksheet
    For Each sht In shtCol
        Dim pageCnt As Long: pageCnt = CheckSheetExistRed(sht)
        col.add sht.name & " : " & pageCnt
        allCnt = allCnt + pageCnt
    Next
    
    col.add "all : " & allCnt
    Dim temp As Variant: temp = ColToAry(col)
    MsgBox Join(temp, vbCrLf)
    
    Finalize
End Sub

'1�V�[�g������Ԏ����܂ރy�[�W�����y�[�W���邩
Private Function GetSheetRedPageCnt(ByVal sht As Worksheet) As Long
    Dim eachPageList As Collection
    Set eachPageList = GetEachPageArea(sht)
    
    Dim cnt As Long: cnt = 0
    Dim r As Range
    For Each r In eachPageList
        If CheckRangeFontColor(r, CLR_RED, False) Then
            cnt = cnt + 1
        End If
    Next
    
    GetSheetRedPageCnt = cnt
End Function

'*****************************************
'2
'*****************************************
'�C���y�[�W�W�v
Public Function GetModPageInfo(ByVal sht As Worksheet) As Long
    
    '����y�[�W�e�ŏI�s�̃��X�g���擾
    Dim hBreakList As Collection
    Set hBreakList = GetPageBreakRowList(sht)
        
    '�C���y�[�W��
    Dim modPageCnt As Long: modPageCnt = 0
        
    '����y�[�W�e�J�n�s(�J�E���g�Ώۂ�5�s�ڂ���B4�s�ڂ܂Ńw�b�_�[)
    Dim sttRow As Long: sttRow = 5
    
    Dim endRow As Variant
    For Each endRow In hBreakList

        '���y�[�W�s���܂łɎw���ɒl�����邩
        Dim chkCol As Integer: chkCol = 2
        Dim r As Range: Set r = sht.Range(sht.Cells(sttRow, chkCol), sht.Cells(endRow, chkCol))
        If WorksheetFunction.CountA(r) > 0 Then
            modPageCnt = modPageCnt + 1
        End If
        
        '���y�[�W�`���s��
        sttRow = endRow + 1
    Next
    
    GetModPageInfo = modPageCnt
End Function

'*****************************************
'����
'*****************************************
'�����F�Ɏw��F���܂܂�邩�`�F�b�N����(�Z���͈�)
'param area �Z���͈�
'param colorval �F
'param containBlank �󔒃Z���̐F���m�F���邩
Private Function CheckRangeFontColor(ByVal area As Range, ByVal colorVal As Long, ByVal containBlank As Boolean) As Boolean
    Dim r As Range
    For Each r In area
        '�w��͈͓��̈ꌏ�ł�����ΏI��
        If CheckCellFontColor(r, colorVal, containBlank) Then
            CheckRangeFontColor = True
            Exit Function
        End If
    Next
    
    CheckRangeFontColor = False
End Function

'�����F�Ɏw��F���܂܂�邩�`�F�b�N����(�P�̃Z��)
'param r �Z��1��
'param colorval �F
'param containBlank �󔒃Z���̐F���m�F���邩
Private Function CheckCellFontColor(ByVal r As Range, ByVal colorVal As Long, ByVal containBlank As Boolean) As Boolean
    
    '�󔒃Z�����`�F�b�N���Ȃ��ꍇ
    If r.Value = "" And Not containBlank Then
        CheckCellFontColor = False
        Exit Function
    End If
    
    '�Z�����̂̐F
    If r.Font.color = colorVal Then
        CheckCellFontColor = True
        Exit Function
    End If
    
    '������ȊO�͏I��
    If VarType(r.Value) <> vbString Then
        CheckCellFontColor = False
        Exit Function
    End If
    
    '�ꕶ���P�ʂ̐F
    If r.Value <> "" And Not r.HasFormula Then
        Dim i As Integer
        For i = 1 To r.Characters.Count
            If r.Characters(i, 1).Font.color = colorVal Then
                CheckCellFontColor = True
                Exit Function
            End If
        Next
    End If
    
    CheckCellFontColor = False
End Function


