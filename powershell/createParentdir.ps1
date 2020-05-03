# ===========================================
# フォルダ名に親フォルダ(_を冒頭に追加)作成
# ===========================================
# -------------------------------------------
# 定数
# -------------------------------------------
# $TARGET_DIR = "C:\_git\__worktmp\"
$TARGET_DIR = "C:\_git\__worktmp\______temp\"

# -------------------------------------------
# 関数
# -------------------------------------------
# 共通部品
. .\common\module.ps1

function main {
    # フォルダ一覧取得 再帰なし
    Get-ChildItem -Path $TARGET_DIR -Directory | ForEach-Object {
        $newDir = (Split-Path $_.FullName -Parent) + "/_" + $_.Name
        # フォルダ作成
        New-Item $newDir -ItemType Directory -Force
        # 新しいフォルダに移動
        Move-Item -Path $_.FullName -Destination ($newDir + "/" + $_.Name)
        Write-ErrorLog $TARGET_DIR
    }
}
# 配下にフォルダが無い時のみ
function main2 {
    # 失敗したので親フォルダを削除したい
    # 1. Everythingでフォルダ配下のファイル・フォルダ一覧を出す
    #     aaaaa1
    #     aaaaa2
    #     aaaaa1/bbbbb1
    #     aaaaa2/bbbbb2
    #     aaaaa1/bbbbb1/cccc1.txt
    #     aaaaa1/bbbbb1/cccc.2txt
    # 2. 子フォルダのみを選択し、別の場所へ切り取り貼り付け
    #     aaaaa1/bbbbb1
    #     aaaaa2/bbbbb2
    #     これらをxxxx1に移動
    #     xxxxx1/bbbbb1
    #     xxxxx1/bbbbb2

    # フォルダ一覧取得 再帰なし
    Get-ChildItem -Path $TARGET_DIR -Directory | ForEach-Object {
        # 配下にフォルダがない場合
        if ((Get-ChildItem -Path $_.FullName -Directory | Measure-Object).Count -eq 0) {
            Write-Host $_.FullName

            $newDir = (Split-Path $_.FullName -Parent) + "/_" + $_.Name
            # フォルダ作成
            New-Item $newDir -ItemType Directory -Force
            # 新しいフォルダに移動
            Move-Item -Path $_.FullName -Destination ($newDir + "/" + $_.Name)
        }
        Write-ErrorLog $TARGET_DIR
    }
}

# -------------------------------------------
# 実行
# -------------------------------------------
main
# main2