# ===========================================
# 同名の親フォルダを作成する
#
# 【使い方】
# .\createParentdir.ps1 -workpath C:\_git\__worktmp\______temp\
# 
# 【失敗したので親フォルダを削除したい】
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
# ===========================================
# -------------------------------------------
# パラメータ
# -------------------------------------------
Param(
    [ValidateSet("C:\_git\__worktmp\" , "C:\_git\__worktmp\______temp\")]$workpath #作業フォルダ
)

# -------------------------------------------
# 関数
# -------------------------------------------
# 共通部品
. .\common\module.ps1

function main {
    param (
        $workpath_
    )
    # フォルダ一覧取得 再帰なし
    Get-ChildItem -Path $workpath_ -Directory | ForEach-Object {
        $newDir = (Split-Path $_.FullName -Parent) + "/_" + $_.Name
        # フォルダ作成
        New-Item $newDir -ItemType Directory -Force
        # 新しいフォルダに移動
        Move-Item -Path $_.FullName -Destination ($newDir + "/" + $_.Name)
        Write-ErrorLog $workpath_
    }
}

# -------------------------------------------
# 実行
# -------------------------------------------
# パラメータ指定がない場合
if ($null -eq $workpath) {
    $workpath = "C:\_git\__worktmp\" 
}
main $workpath