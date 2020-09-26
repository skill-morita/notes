Attribute VB_Name = "UtilDriver"
'*******************************************************
'UtilDriver
'Note:  登録用マクロ
'*******************************************************

'@Folder("VBAProject")
Option Explicit

'フッター設定
Public Sub test()
    Call SetPrintFooter(ActiveSheet, "", "&P/&Nページ", "&F")
End Sub

'ページ設定
Public Sub test2()

    With ActiveWorkbook
        .Worksheets.Select
        Call SetPageBreakPreview
        Call SetZoom(80)
        .Worksheets(1).Select
    End With
End Sub

'選択範囲の空白セルを書式リセット
Public Sub SetClearFormatBlankCells()
    Initialize
    Dim r As Range: Set r = FilterBlankCells(Selection)
    Call SetClearOnlyFormat(r)
    Finalize
End Sub

'-----------------------------------------
'一般
'-----------------------------------------
Public Sub カーソルA1移動()
    Call SetActiveA1
End Sub
'-----------------------------------------
'書式
'-----------------------------------------
Public Sub 縮小表示OFF()
    ' セルの文字列をセル幅に合わせてすべて表示する(リボンでは設定できない)
    Call SetShrinkToFit(ActiveCell, False)
End Sub
Public Sub 縮小表示ON()
    Call SetShrinkToFit(ActiveCell, True)
End Sub

'-----------------------------------------
'色
'-----------------------------------------
Public Sub 文字黒()
    Call SetFontBlack
End Sub

Public Sub 文字赤()
    Call SetFontRed
End Sub

Public Sub 文字青()
    Call SetFontBlue
End Sub

Public Sub 背景C()
    Call SetBackClear
End Sub

Public Sub 背景赤()
    Call SetBackRed
End Sub

Public Sub 背景青()
    Call SetBackBlue
End Sub

Public Sub 背景黄()
    Call SetBackYellow
End Sub

Public Sub 背景灰()
    Call SetBackGray
End Sub

'-----------------------------------------
'データ取得
'-----------------------------------------
Public Sub GetSelectCell()
    Dim dic As Object
    Set dic = GetSelectCellInfo(ActiveCell)
    If dic Is Nothing Then
        MsgBox "セルが選択されていません"
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
        
        '取得
        Dim dic As Object
        Set dic = GetSelectShapeInfo(shp)
        If dic Is Nothing Then
            MsgBox "図形が選択されていません"
            Exit Sub
        End If
        
        '表示
        Dim key As Variant
        For Each key In dic.Keys
            Debug.Print "Key: " & key, "Value: " & dic.Item(key)
        Next
        
        dic = Nothing
    Next
End Sub

'-----------------------------------------
'グループ化でアウトライン
'-----------------------------------------
Public Sub SetGroupByMethod()
    Dim sht As Worksheet: Set sht = ActiveSheet
    '初期化
    Call SetClearGroup(sht.UsedRange)
    
    '----------------------
    'グループ化の区切りを検索
    '----------------------
    '検索書式設定
    Call SetFindFormatClear
    Call SetFindFormatBackColor(CLR_SKYBLUE)
    'P列に背景色水色の「メソッド名」がある
    Dim colMethodAddList As Collection
    Set colMethodAddList = SetFind("P3:P" & GetEndRow(sht), "メソッド名", True)
    '検索書式初期化
    Call SetFindFormatClear
    '無い場合は終了
    If colMethodAddList Is Nothing Then
        Exit Sub
    End If
    
    '----------------------
    'グループ設定
    '----------------------
    Dim i As Integer
    For i = 1 To colMethodAddList.Count
        'メソッド定義行 1行目
        Dim selCellAdd As String
        selCellAdd = colMethodAddList.Item(i)
        '背景色水色でなくなる場所まで下りる
        Dim tmpR As Range: Set r = sht.Range(selCellAdd)
        Do
            '一つ下の行に下りる
            Set tmpR = tmpR.Offset(1, 0)
        Loop Until tmpR.Interior.color <> CLR_SKYBLUE
        
        'グループ化開始行
        Dim grpSttCell As Range: Set grpSttCell = tmpR
        'グループ化最終行
        Dim grpEndCell As Range
        If i = colMethodAddList.Count Then
            '最後の定義の場合
            Set grpEndCell = sht.Range("P" & GetEndRow(sht))
        Else
            '次の定義の一行前
            Set grpEndCell = sht.Range(colMethodAddList.Item(i + 1)).Offset(-1, 0)
        End If
        
        'グループ化
        Call SetGroup(sht.Range(grpSttCell, grpEndCell))
    Next
End Sub

Public Sub 横G全閉()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim sht As Worksheet
    For Each sht In wb.Sheets
        Call SetHideGroup(sht, True)
    Next
End Sub

Public Sub 横G全開()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim sht As Worksheet
    For Each sht In wb.Sheets
        Call SetOpenGroup(sht, True)
    Next
End Sub

