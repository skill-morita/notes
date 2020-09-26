Attribute VB_Name = "UtilSheetIn"
'*******************************************************
'UtilSheetIn
'Note:  シート内の操作
'*******************************************************
Option Explicit

'*****************************************
'判定
'*****************************************
'*****************************************
'取得
'*****************************************
'空白セルに絞り込み
Public Function FilterBlankCells(ByRef r As Range) As Range
On Error GoTo ErrHandl
    Set FilterBlankCells = r.SpecialCells(xlCellTypeBlanks)
    Exit Function
ErrHandl:
    '空白セルが無い時エラー発生
    Set FilterBlankCells = Nothing
End Function

'セル情報取得
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
    'セルが無い時エラー発生
    Set GetSelectCellInfo = Nothing
End Function

'シェイプ情報取得
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
    'セルが無い時エラー発生
    Set GetSelectShapeInfo = Nothing
End Function
'*****************************************
'設定
'*****************************************
'全てのシートのカーソルをA1に移動
Public Sub SetActiveA1()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim sht As Worksheet
    For Each sht In wb.Sheets
        sht.Activate
        sht.Range("A1").Select
    Next sht
    '最初に移動
    wb.Sheets(1).Activate
End Sub
'-----------------------------------------
'書式
'-----------------------------------------
Public Sub SetShrinkToFit(ByRef r As Range, ByVal flg As Boolean)
    'セルの文字列をセル幅に合わせてすべて表示する
    r.ShrinkToFit = flg
End Sub

'-----------------------------------------
'クリア
'-----------------------------------------
Public Sub SetClearOnlyFormat(ByRef r As Range)
    r.ClearFormats
End Sub
Public Sub SetClearOnlyVal(ByRef r As Range)
    r.ClearContents
End Sub
'-----------------------------------------
'表示
'-----------------------------------------
'標準
Public Sub SetNormalView()
    ActiveWindow.View = xlNormalView
End Sub
'改ページプレビュー
Public Sub SetPageBreakPreview()
    ActiveWindow.View = xlPageBreakPreview
End Sub
'レイアウトビュー
Public Sub SetPageLayoutView()
    ActiveWindow.View = xlPageLayoutView
End Sub
'-----------------------------------------
'ズーム
'-----------------------------------------
Public Sub SetZoom(ByVal zoomPer As Integer)
    ActiveWindow.Zoom = zoomPer
End Sub
'-----------------------------------------
'グループ化/アウトライン
'-----------------------------------------
'グループ化初期設定
Public Sub SetInitGroup(ByVal sht As Worksheet)
    With sht.Outline
        .AutomaticStyles = False
        '.SummaryRow = xlSummaryBelow        '行のグループ化マークの+/-を下側に(標準設定)
        '.SummaryColumn = xlSummaryOnRight   '列のグループ化マークの+/-を右側に(標準設定)
        .SummaryRow = xlSummaryAbove        '行のグループ化マークの+/-を上側に
        .SummaryColumn = xlSummaryOnLeft    '列のグループ化マークの+/-を左側に
    End With
End Sub

'グループ化設定
'41L~45Lを選択して設定→40Lにマーク出現。41L~45Lはたたまれる
Public Sub SetGroup(ByVal rng As Range)
    '初期設定
    Call SetInitGroup(rng.Parent)
    rng.Rows.Group
End Sub

'グループ化解除
Public Sub SetClearGroup(ByVal rng As Range)
'グループが無い場合エラー回避
On Error GoTo ErrHandl
    rng.Rows.Ungroup
    'rng.ClearOutline こっちのほうが良い？
ErrHandl:
End Sub

'アウトラインを閉じる
Public Sub SetHideGroup(ByVal sht As Worksheet, ByVal colFlg As Boolean)
    If colFlg Then
        sht.Outline.ShowLevels ColumnLevels:=1
    Else
        sht.Outline.ShowLevels RowLevels:=1
    End If
End Sub

'アウトラインを開く
Public Sub SetOpenGroup(ByVal sht As Worksheet, ByVal colFlg As Boolean)
    If colFlg Then
        sht.Outline.ShowLevels ColumnLevels:=2
    Else
        sht.Outline.ShowLevels RowLevels:=2
    End If
End Sub

