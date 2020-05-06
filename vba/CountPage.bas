Attribute VB_Name = "CountPage"
Option Explicit

'TODO : シェイプの赤文字取得

'*****************************************
'1
'*****************************************
'ブック全体のページ数取得
Public Sub CheckBookExistRed()
    Initialize

    Dim col As New Collection
    Dim allCnt As Long
    
    'シートタブが赤のものを取得
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

'1シートあたり赤字を含むページが何ページあるか
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
'修正ページ集計
Public Function GetModPageInfo(ByVal sht As Worksheet) As Long
    
    '印刷ページ各最終行のリストを取得
    Dim hBreakList As Collection
    Set hBreakList = GetPageBreakRowList(sht)
        
    '修正ページ数
    Dim modPageCnt As Long: modPageCnt = 0
        
    '印刷ページ各開始行(カウント対象は5行目から。4行目までヘッダー)
    Dim sttRow As Long: sttRow = 5
    
    Dim endRow As Variant
    For Each endRow In hBreakList

        '改ページ行くまでに指定列に値があるか
        Dim chkCol As Integer: chkCol = 2
        Dim r As Range: Set r = sht.Range(sht.Cells(sttRow, chkCol), sht.Cells(endRow, chkCol))
        If WorksheetFunction.CountA(r) > 0 Then
            modPageCnt = modPageCnt + 1
        End If
        
        '次ページ冒頭行へ
        sttRow = endRow + 1
    Next
    
    GetModPageInfo = modPageCnt
End Function

'*****************************************
'共通
'*****************************************
'文字色に指定色が含まれるかチェックする(セル範囲)
'param area セル範囲
'param colorval 色
'param containBlank 空白セルの色も確認するか
Private Function CheckRangeFontColor(ByVal area As Range, ByVal colorVal As Long, ByVal containBlank As Boolean) As Boolean
    Dim r As Range
    For Each r In area
        '指定範囲内の一件でもあれば終了
        If CheckCellFontColor(r, colorVal, containBlank) Then
            CheckRangeFontColor = True
            Exit Function
        End If
    Next
    
    CheckRangeFontColor = False
End Function

'文字色に指定色が含まれるかチェックする(単体セル)
'param r セル1つ
'param colorval 色
'param containBlank 空白セルの色も確認するか
Private Function CheckCellFontColor(ByVal r As Range, ByVal colorVal As Long, ByVal containBlank As Boolean) As Boolean
    
    '空白セルをチェックしない場合
    If r.Value = "" And Not containBlank Then
        CheckCellFontColor = False
        Exit Function
    End If
    
    'セル自体の色
    If r.Font.color = colorVal Then
        CheckCellFontColor = True
        Exit Function
    End If
    
    '文字列以外は終了
    If VarType(r.Value) <> vbString Then
        CheckCellFontColor = False
        Exit Function
    End If
    
    '一文字単位の色
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


