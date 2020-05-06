# ===========================================
# ショートカット類を1階層目に移動
# ===========================================
# -------------------------------------------
# 定数
# -------------------------------------------
$TARGET_DIR = "C:\_git\__worktmp\"

# -------------------------------------------
# 関数
# -------------------------------------------
# 共通部品
. .\common\module.ps1

function main {
    # フォルダ一覧取得 再帰なし
    Get-ChildItem -Path $TARGET_DIR -Directory | ForEach-Object {
        # リンクはニコニコ動画、ニコニコ静画、ニコニ立体の何れかを含むもの
        $files = Get-ChildItem -Path $_.FullName -Include "*ニコニ*.url", "*.wav" , "*.zip" -Recurse
        foreach ($item in $files) {
            $item.FullName
            Move-Item $item.FullName -Destination ($_.FullName + "/" + $item.Name)
        }
    }
}

# -------------------------------------------
# 実行
# -------------------------------------------
main