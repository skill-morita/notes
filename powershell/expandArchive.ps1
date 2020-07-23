# ===========================================
# zipファイルを展開し、そのフォルダ内にzipを移動する
#
# 【使い方】
# .\expandArchive.ps1 -workpath C:\_git\__worktmp\
# ===========================================
# -------------------------------------------
# パラメータ
# -------------------------------------------
Param(
    [ValidateSet("C:\_git\__worktmp\" , "C:\_git\__worktmp\______temp\")]$workpath, #作業フォルダ
    [ValidateSet(".zip")]$ext # 拡張子
)
# -------------------------------------------
# 関数
# -------------------------------------------
# 共通部品
.".\common\module.ps1"

function main {
param (
    $workpath_,
    $ext_
)
    # 圧縮ファイル一覧取得 再帰なし
    $zipfile = Get-ChildItem -Path $workpath_ -File -Filter *$ext_
    # $zipfile
    $zipfile | ForEach-Object {
        $dirnm = ($_.FullName).Replace($ext_, "")
        # zip展開
        # TODO zip以外の拡張子
        Expand-Archive -Path $_.FullName -DestinationPath $dirnm 
        Write-ErrorLog $workpath_
        # 展開したフォルダにzipを移動
        Move-Item $_.FullName $dirnm
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
if ($null -eq $ext) {
    $ext = ".zip"
}
main $workpath $ext