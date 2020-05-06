Attribute VB_Name = "UtilOperateFile"
Option Explicit

'フォルダ内のファイル取得
Public Function GetFiles(ByVal FolderPath As String) As Collection
    'ファイルシステムオブジェクト取得
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Folderオブジェクト取得
    Dim f As Object
    Set f = fso.GetFolder(FolderPath)

    'Fileコレクションを返す
    Dim res As New Collection
    Dim itm As Variant
    For Each itm In f.Files
        res.add itm.Path
    Next
    
    Set GetFiles = res
End Function

'指定セルの内容をCSVに出力する
'出力パスを取得
'選択セル内容を取得
'selection.cells(1.1).text
'ファイル出力


'*********************************************
'Excel
'*********************************************
Public Function OpenExcel(ByVal filepath As String, ByVal readOnlyFlg As Boolean) As Workbook
    Dim res As Workbook
    Set res = Workbooks.Open(filepath, readOnlyFlg)
End Function

Public Sub CloseExcel(ByRef wb As Workbook, ByVal saveChangeFlg As Boolean)
    Call wb.Close(saveChangeFlg)
End Sub

'個人マクロブック以外の開いているブックを取得する
Public Function GetOpenBook() As Collection
    Dim res As New Collection
    Dim wb As Workbook
    For Each wb In Workbooks
        If MACRO_BOOK <> wb.name Then
            res.add wb
        End If
    Next
    Set GetOpenBook = res
End Function

