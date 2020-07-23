# ===========================================
# フォルダ作成
#
# 【使い方】
# .\createDirs.ps1 -workpath C:\_git\__worktmp\
# ===========================================
# -------------------------------------------
# パラメータ
# -------------------------------------------
Param(
    [ValidateSet("C:\_git\__worktmp\" , "C:\_git\__worktmp\______temp\")]$workpath #作業フォルダ
)

$DIR_NAME_LIST = @(
    "__金星のダンス"
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
    $DIR_NAME_LIST | ForEach-Object {
        New-Item -Path ($workpath_ + "/" + $_) -ItemType Directory
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