# ===========================================
# 配下の各フォルダにファイル配布(フォルダを指定した場合は配下のファイルもコピー、配布先not再帰)
# 
# 【使い方】
# .\distributeFile.ps1 -workpath C:\_git\__worktmp\______temp\
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
# 共通部品(エラー出るときはps実行フォルダを確かめる？)
. .\common\module.ps1

# コピーファイルパス
[string]$srcFilepath = (Read-Host コピーしたいファイルパス)
function main {
    param (
        $workpath_
    )

    # フォルダ一覧取得 再帰なし
    $dirList = Get-ChildItem -Path $workpath_ -Directory

    $dirList | ForEach-Object {
        # ファイルコピー
        $_.FullName
        Copy-Item -Path $srcFilepath $_.FullName -Recurse
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