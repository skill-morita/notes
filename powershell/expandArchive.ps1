# ===========================================
# zipファイルを展開し、そのフォルダ内にzipを移動する
# ===========================================
# -------------------------------------------
# 定数
# -------------------------------------------
# 処理フォルダ
# $TARGET_DIR = "C:\_git\__worktmp\"
$TARGET_DIR = "C:\_git\__worktmp\______temp\"
# 拡張子
$EXT = ".zip"

# -------------------------------------------
# 関数
# -------------------------------------------
# 共通部品
.".\common\module.ps1"

function main {
    # 圧縮ファイル一覧取得 再帰なし
    $zipfile = Get-ChildItem -Path $TARGET_DIR -File -Filter *$EXT
    # $zipfile
    $zipfile | ForEach-Object {
        $dirnm = ($_.FullName).Replace($EXT, "")
        # zip展開
        Expand-Archive -Path $_.FullName -DestinationPath $dirnm 
        Write-ErrorLog $TARGET_DIR
        # 展開したフォルダにzipを移動
        Move-Item $_.FullName $dirnm
        Write-ErrorLog $TARGET_DIR
    }
}

# -------------------------------------------
# 実行
# -------------------------------------------
main