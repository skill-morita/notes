# ===========================================
# フォルダ名に親フォルダ(_を冒頭に追加)作成
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
    $dirList = Get-ChildItem -Path $TARGET_DIR -Directory

    $dirList | ForEach-Object {
        $newDir = (Split-Path $_.FullName -Parent) + "/_" + $_.Name
        # フォルダ作成
        New-Item $newDir -ItemType Directory -Force
        # 新しいフォルダに移動
        Move-Item -Path $_.FullName -Destination ($newDir + "/" + $_.Name)
        Write-ErrorLog
    }
}

# -------------------------------------------
# 実行
# -------------------------------------------
main