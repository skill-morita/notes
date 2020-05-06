Attribute VB_Name = "UtilPrint"
'*******************************************************
'UtilPrint
'Note:  ����֘A
'*******************************************************

Option Explicit

'=========================================
'����
'=========================================

'=========================================
'�擾
'=========================================
'����y�[�W��
Public Function GetPrintPageCnt(ByVal sht As Worksheet) As Long
    GetPrintPageCnt = sht.PageSetup.Pages.Count
End Function

'����͈͎擾
Public Function GetPrintArea(ByVal sht As Worksheet) As Range
    Set GetPrintArea = sht.Range(sht.PageSetup.PrintArea)
End Function

'���1�y�[�W���̃Z���͈͎擾
Public Function GetEachPageArea(ByVal sht As Worksheet) As Collection
    '�e�y�[�W�ŏI�s���X�g���擾
    Dim hPageList As Collection
    Set hPageList = GetPageBreakRowList(sht)
    
    '����̗�͈͂��擾
    Dim sttCol As Long: sttCol = GetSttCol(sht)
    Dim endCol As Long: endCol = GetEndCol(sht)
    
    Dim res As New Collection
    
    Dim i As Long
    Dim sttRow As Long: sttRow = GetSttRow(sht)
    Dim endRow As Long
    For i = 1 To hPageList.Count
        endRow = hPageList(i)
        Dim r As Range: Set r = sht.Range(sht.Cells(sttRow, sttCol), sht.Cells(endRow, endCol))
        Call res.add(r)
        sttRow = endRow + 1
    Next
    
    Set GetEachPageArea = res
End Function

'�������y�[�W�̊e�ŏI�s���擾
Public Function GetPageBreakRowList(ByVal sht As Worksheet) As Collection
    Dim res As New Collection
    
    Dim pageBreak As HPageBreak
    For Each pageBreak In sht.HPageBreaks
        'hBreak�����͉̂��y�[�W�J�n�s
        Call res.add(pageBreak.Location.row - 1)
    Next
    
    '�ŏI�s�ǉ�
    Call res.add(GetEndRow(sht))

    Set GetPageBreakRowList = res
End Function

'-----------------------------------------
'����s
'-----------------------------------------
'����J�n�s�擾
Public Function GetSttRow(ByVal sht As Worksheet) As Long
    GetSttRow = GetPrintArea(sht).Rows.row
End Function

'����ŏI�s�擾
Public Function GetEndRow(ByVal sht As Worksheet) As Long
    Dim r As Range: Set r = GetPrintArea(sht)
    GetEndRow = r.Rows.row + r.Rows.Count - 1
End Function

'-----------------------------------------
'�����
'-----------------------------------------
'����J�n��擾
Public Function GetSttCol(ByVal sht As Worksheet) As Long
    GetSttCol = GetPrintArea(sht).Columns.Column
End Function

'����ŏI��擾
Public Function GetEndCol(ByVal sht As Worksheet) As Long
    Dim r As Range: Set r = GetPrintArea(sht)
    GetEndCol = r.Columns.Column + r.Columns.Count - 1
End Function

'=========================================
'�ݒ�
'=========================================
'-----------------------------------------
'�g��k��
'-----------------------------------------
'����y�[�W�ݒ�
'param per �{��(100��100%�A0�͎���)
Public Sub SetPrintZoom(ByRef sht As Worksheet, ByVal percnt As Integer)
    Application.PrintCommunication = False
    With sht.PageSetup
        If percnt = 0 Then
            .Zoom = False
        Else
            .Zoom = percnt
        End If
    End With
    Application.PrintCommunication = True
End Sub

'����y�[�W�ݒ�
'param wPageCnt �������y�[�W�Ɏ��߂邩(0�͎���)
'param hPageCnt �c�����y�[�W�Ɏ��߂邩(0�͎���)
Public Sub SetPrintFitToPage(ByRef sht As Worksheet, ByVal wPageCnt As Integer, ByVal hPageCnt As Integer)
    Application.PrintCommunication = False
    With sht.PageSetup
        If wPageCnt = 0 Then
            .FitToPagesWide = False
        Else
            .FitToPagesWide = wPageCnt
        End If
        
        If hPageCnt = 0 Then
            .FitToPagesTall = False
        Else
            .FitToPagesTall = hPageCnt
        End If
    End With
    Application.PrintCommunication = True
End Sub

'-----------------------------------------
'�w�b�_�[�t�b�^�[
'-----------------------------------------
Public Sub SetPrintHeader(ByRef sht As Worksheet, _
                            ByVal left As String, _
                            ByVal center As String, _
                            ByVal right As String)
    With sht.PageSetup
        .LeftHeader = left
        .CenterHeader = center
        .RightHeader = right
    End With
End Sub

Public Sub SetPrintFooter(ByRef sht As Worksheet, _
                            ByVal left As String, _
                            ByVal center As String, _
                            ByVal right As String)
    With sht.PageSetup
        .LeftFooter = left
        .CenterFooter = center
        .RightFooter = right
    End With
End Sub

