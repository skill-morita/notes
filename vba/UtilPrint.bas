Attribute VB_Name = "UtilPrint"
'*******************************************************
'UtilPrint
'Note:  印刷関連
'*******************************************************

Option Explicit

'=========================================
'判定
'=========================================

'=========================================
'取得
'=========================================
'印刷ページ数
Public Function GetPrintPageCnt(ByVal sht As Worksheet) As Long
    GetPrintPageCnt = sht.PageSetup.Pages.Count
End Function

'印刷範囲取得
Public Function GetPrintArea(ByVal sht As Worksheet) As Range
    Set GetPrintArea = sht.Range(sht.PageSetup.PrintArea)
End Function

'印刷1ページごのセル範囲取得
Public Function GetEachPageArea(ByVal sht As Worksheet) As Collection
    '各ページ最終行リストを取得
    Dim hPageList As Collection
    Set hPageList = GetPageBreakRowList(sht)
    
    '印刷の列範囲を取得
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

'水平改ページの各最終行を取得
Public Function GetPageBreakRowList(ByVal sht As Worksheet) As Collection
    Dim res As New Collection
    
    Dim pageBreak As HPageBreak
    For Each pageBreak In sht.HPageBreaks
        'hBreakが持つのは改ページ開始行
        Call res.add(pageBreak.Location.row - 1)
    Next
    
    '最終行追加
    Call res.add(GetEndRow(sht))

    Set GetPageBreakRowList = res
End Function

'-----------------------------------------
'印刷行
'-----------------------------------------
'印刷開始行取得
Public Function GetSttRow(ByVal sht As Worksheet) As Long
    GetSttRow = GetPrintArea(sht).Rows.row
End Function

'印刷最終行取得
Public Function GetEndRow(ByVal sht As Worksheet) As Long
    Dim r As Range: Set r = GetPrintArea(sht)
    GetEndRow = r.Rows.row + r.Rows.Count - 1
End Function

'-----------------------------------------
'印刷列
'-----------------------------------------
'印刷開始列取得
Public Function GetSttCol(ByVal sht As Worksheet) As Long
    GetSttCol = GetPrintArea(sht).Columns.Column
End Function

'印刷最終列取得
Public Function GetEndCol(ByVal sht As Worksheet) As Long
    Dim r As Range: Set r = GetPrintArea(sht)
    GetEndCol = r.Columns.Column + r.Columns.Count - 1
End Function

'=========================================
'設定
'=========================================
'-----------------------------------------
'拡大縮小
'-----------------------------------------
'印刷ページ設定
'param per 倍率(100→100%、0は自動)
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

'印刷ページ設定
'param wPageCnt 横を何ページに収めるか(0は自動)
'param hPageCnt 縦を何ページに収めるか(0は自動)
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
'ヘッダーフッター
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

