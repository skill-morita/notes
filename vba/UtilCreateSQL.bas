Attribute VB_Name = "UtilCreateSQL"
'*******************************************************
'UtilCreateSQL
'Note:  SQL�쐬
'*******************************************************
Option Explicit

Public Sub SQL�쐬()
    Call Initialize
    '�擾�Ώ� �ԃV�[�g
    Dim shtList As Collection: Set shtList = GetColorSheet(CLR_RED)
    
    '�o�� ���V�[�g
    Dim tmp As Collection: Set tmp = GetColorSheet(CLR_YELLOW)
    Dim outSht As Worksheet: Set outSht = tmp.Item(1)
    Call InitOutSheet(outSht)
    
    Dim sht As Worksheet
    For Each sht In shtList
        '�e�[�u����
        Dim tblnm As String: tblnm = sht.Range("B2").Value
        'PK��ʒu
        Const pkColCnt As Long = 3
        '����s�ʒu
        Const cnstRowCnt As Long = 3
        '�^�s�ʒu
        Const fmtRowCnt As Long = 4
        '�J�����s�ʒu
        Const keyRowCnt As Long = 6
        
        '�e�V�[�g��SQL�o��
        Dim dic As Dictionary
        Set dic = CreateSqls(sht, tblnm, pkColCnt, keyRowCnt, fmtRowCnt, cnstRowCnt)
        
        '�܂Ƃ߃V�[�g�ɏo��
        Call OutputSqls(outSht, dic)
        
        Set sht = Nothing
    Next
    
    '�t�B���^�ݒ�
    outSht.UsedRange.AutoFilter
    
    MsgBox "End"
    
    Call Finalize
End Sub

Private Sub InitOutSheet(ByRef sht As Worksheet)
    With sht
        '�t�B���^����
        .AutoFilterMode = False
        '�f�[�^���Z�b�g
        .Cells.ClearFormats
        
        '�w�b�_
        .Cells(1, 1).Value = "No"
        .Cells(1, 2).Value = "Sht"
        .Cells(1, 3).Value = "Ptn"
        .Cells(1, 4).Value = "Typ"
        .Cells(1, 5).Value = "--Sql"
    End With
End Sub

'SQL�������V�[�g�ɏo��
Private Function CreateSqls(ByVal sht As Worksheet, ByVal tblnm As String, _
                            ByVal pkColCnt As Long, ByVal keyRowCnt As Long, _
                            ByVal fmtRowCnt As Long, ByVal cnstRowCnt As Long) As String
    Const ptnColCnt As Long = 1
    Dim sqlAry As Variant
    sqlAry = Array(DEL, INS, UPD, SEL)
    '�J��
    Call SetOpenGroup(sht, True)
    '�f�[�^�s�͈�
    Dim sttRow As Long: sttRow = keyRowCnt + 1
    Dim endRow As Long: endRow = GetEndFillRow(sht, ptnColCnt)
    '�f�[�^��͈�
    Dim sttCol As Long: sttCol = pkColCnt
    Dim endCol As Long: endCol = GetEndFillColumn(sht, fmtRowCnt)
    '�o�͗�ʒu
    Dim outContent As Long: outContent = endCol + 1
    'PK��
    Dim idNm As String: idNm = sht.Cells(keyRowCnt, pkColCnt).Value
    
    '�J�������͈�
    Dim keyRng As Range: Set keyRng = sht.Range(sht.Cells(keyRowCnt, sttCol), sht.Cells(keyRowCnt, endCol))
    '����(�񖼍s���Ƌ��J�E���g�ɓ����Ă��܂��̂�)
    sht.Cells(keyRowCnt + 1, outContent).Value = Now()
    'SQL�쐬
    Dim dic As Dictionary: Set dic = New Dictionary
    Dim i As Long
    For i = sttRow To endRow
        Dim idVal As String: idVal = sht.Cells(i, pkColCnt).Value
        'TODO
    Next
    
    'TODO�@�^�g������
End Function

Private Function OutputSqls() As String


End Function

Public Function GetEndFillRow(ByVal sht As Worksheet, ByVal cCnt As Long) As Long
    GetEndFillRow = sht.Cells(sht.Rows.Count, cCnt).End(xlUp).row
End Function
Public Function GetEndFillColumn(ByVal sht As Worksheet, ByVal rCnt As Long) As Long
    GetEndFillColumn = sht.Cells(rCnt, sht.Columns.Count).End(xlToLeft).Column
End Function
