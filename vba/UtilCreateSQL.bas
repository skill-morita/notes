Attribute VB_Name = "UtilCreateSQL"
'*******************************************************
'UtilCreateSQL
'Note:  SQL作成
'*******************************************************
Option Explicit

Public Sub SQL作成()
    Call Initialize
    '取得対象 赤シート
    Dim shtList As Collection: Set shtList = GetColorSheet(CLR_RED)
    
    '出力 黄シート
    Dim tmp As Collection: Set tmp = GetColorSheet(CLR_YELLOW)
    Dim outSht As Worksheet: Set outSht = tmp.Item(1)
    Call InitOutSheet(outSht)
    
    Dim sht As Worksheet
    For Each sht In shtList
        'テーブル名
        Dim tblnm As String: tblnm = sht.Range("B2").Value
        'PK列位置
        Const pkColCnt As Long = 3
        '制約行位置
        Const cnstRowCnt As Long = 3
        '型行位置
        Const fmtRowCnt As Long = 4
        'カラム行位置
        Const keyRowCnt As Long = 6
        
        '各シートにSQL出力
        Dim dic As Dictionary
        Set dic = CreateSqls(sht, tblnm, pkColCnt, keyRowCnt, fmtRowCnt, cnstRowCnt)
        
        'まとめシートに出力
        Call OutputSqls(outSht, dic)
        
        Set sht = Nothing
    Next
    
    'フィルタ設定
    outSht.UsedRange.AutoFilter
    
    MsgBox "End"
    
    Call Finalize
End Sub

Private Sub InitOutSheet(ByRef sht As Worksheet)
    With sht
        'フィルタ解除
        .AutoFilterMode = False
        'データリセット
        .Cells.ClearFormats
        
        'ヘッダ
        .Cells(1, 1).Value = "No"
        .Cells(1, 2).Value = "Sht"
        .Cells(1, 3).Value = "Ptn"
        .Cells(1, 4).Value = "Typ"
        .Cells(1, 5).Value = "--Sql"
    End With
End Sub

'SQLを書くシートに出力
Private Function CreateSqls(ByVal sht As Worksheet, ByVal tblnm As String, _
                            ByVal pkColCnt As Long, ByVal keyRowCnt As Long, _
                            ByVal fmtRowCnt As Long, ByVal cnstRowCnt As Long) As String
    Const ptnColCnt As Long = 1
    Dim sqlAry As Variant
    sqlAry = Array(DEL, INS, UPD, SEL)
    '開く
    Call SetOpenGroup(sht, True)
    'データ行範囲
    Dim sttRow As Long: sttRow = keyRowCnt + 1
    Dim endRow As Long: endRow = GetEndFillRow(sht, ptnColCnt)
    'データ列範囲
    Dim sttCol As Long: sttCol = pkColCnt
    Dim endCol As Long: endCol = GetEndFillColumn(sht, fmtRowCnt)
    '出力列位置
    Dim outContent As Long: outContent = endCol + 1
    'PK列
    Dim idNm As String: idNm = sht.Cells(keyRowCnt, pkColCnt).Value
    
    'カラム名範囲
    Dim keyRng As Range: Set keyRng = sht.Range(sht.Cells(keyRowCnt, sttCol), sht.Cells(keyRowCnt, endCol))
    '日時(列名行だと橋カウントに入ってしまうので)
    sht.Cells(keyRowCnt + 1, outContent).Value = Now()
    'SQL作成
    Dim dic As Dictionary: Set dic = New Dictionary
    Dim i As Long
    For i = sttRow To endRow
        Dim idVal As String: idVal = sht.Cells(i, pkColCnt).Value
        'TODO
    Next
    
    'TODO　型使いたい
End Function

Private Function OutputSqls() As String


End Function

Public Function GetEndFillRow(ByVal sht As Worksheet, ByVal cCnt As Long) As Long
    GetEndFillRow = sht.Cells(sht.Rows.Count, cCnt).End(xlUp).row
End Function
Public Function GetEndFillColumn(ByVal sht As Worksheet, ByVal rCnt As Long) As Long
    GetEndFillColumn = sht.Cells(rCnt, sht.Columns.Count).End(xlToLeft).Column
End Function
