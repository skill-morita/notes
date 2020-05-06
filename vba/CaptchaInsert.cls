VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaptchaInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Type ImgInfo
    filepath As String
    Per As Integer
        
    left As Single
    Top As Single
    
    Height As Single
    Width As Single
End Type

Public Sub CaptchaInsert()
    Const FolderPath As String = "C:\temp\"
    Const Ext As String = "png"
    Dim interval As Single
    interval = 5

    ' 指定フォルダ配下のファイルリスト取得
    Dim pathList As Collection
    Set pathList = GetFiles(FolderPath)
    
    ' 現在位置取得
    Dim posL As Single
    Dim posT As Single
    posL = Selection.left
    posT = Selection.Top
    
    For Each p In pathList
        ' 拡張子がファイルパスに含まれている
        If InStr(p.Path, Ext) > 0 Then
            Dim iInfo As ImgInfo
            iInfo.filepath = p.Path
            iInfo.left = posL
            iInfo.Top = posT
            iInfo.Per = 100
            
            ' 画像ファイルの挿入
            iInfo = InsertImg(ActiveSheet, iInfo)
            ' 画像のサイズ分下に移動する
            posT = posT + iInfo.Top + interval
            
            ' 挿入した画像は不要なので削除
            Kill p.Path
        End If
    Next
 
End Sub


'画像を挿入する
Private Function InsertImg(ByRef sht As Worksheet, ByRef iInfo As ImgInfo) As ImgInfo
    Dim shp As Shape
    
    ' 画像貼り付け
    With sht
        Set shp = .Shapes.AddPicture( _
                Filename:=iInfo.filepath, _
                linktofile:=False, _
                savewithdocument:=True, _
                left:=iInfo.left, _
                Top:=iInfo.Top, _
                Width:=0, _
                Height:=0)
    End With

    With shp
        ' 一旦縦横のサイズを100%にリセットする
        .LockAspectRatio = False
        .ScaleHeight 1, msoTrue, msoScaleFromTopLeft
        .ScaleWidth 1, msoTrue, msoScaleFromTopLeft

        ' 縦横比固定で指定％に設定
        .LockAspectRatio = True
        .ScaleHeight (iInfo.Per / 100), msoTrue, msoScaleFromTopLeft
        
        ' 最前面に配置
        .ZOrder msoBringToFront
        
        ' 画像サイズ取得
        iInfo.Height = .Height
        iInfo.Width = .Width
    End With

    Set shp = Nothing
    
    ' 画像サイズ返す
    InsertImg = iInfo
End Function



