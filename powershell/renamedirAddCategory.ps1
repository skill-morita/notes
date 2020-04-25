# ===========================================
# フォルダ名にカテゴリ追加
# _カテゴリ_カテゴリ_元のフォルダ名
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

function main ([string]$dirPath, [string]$oldnm, [string]$addStr){
    # フォルダ名変更
    $newnm = ""

    if (($oldnm).Substring(0, 1) -eq "_") {
        $newnm = $addStr + $oldnm
    } else {
        $newnm = $addStr + "_" + $oldnm
    }

    Rename-Item -Path $dirPath -NewName $newnm
    Write-ErrorLog
}

function inputCondition {
    # カテゴリ追加する場合↓----------------
    # 配布方法
    [int]$category = (Read-Host カテゴリ:1 室内、2 室外、3 抽象／4 洋風、5 和風、6 中華、7 現代、8 植物、9 シンプル、10 かわいい、11 きれい／101 スカイドーム)
    [string]$categorynm = ""

    switch ($category) {
        1 {$categorynm = "室内"; break}
        2 {$categorynm = "室外"; break}
        3 {$categorynm = "抽象"; break}
        4 {$categorynm = "洋風"; break}
        5 {$categorynm = "和風"; break}
        6 {$categorynm = "中華"; break}
        7 {$categorynm = "現代"; break}
        8 {$categorynm = "植物"; break}
        9 {$categorynm = "シンプル"; break}
        10 {$categorynm = "かわいい"; break}
        11 {$categorynm = "きれい"; break}
        101 {$categorynm = "skydome"; break}
    }

    # 追加用の空配列
    $addAry = @()
    $addAry += $categorynm
    # カテゴリ追加する場合↑----------------

    # カテゴリ連結
    return "_" + ($addAry -join "_")
}

# -------------------------------------------
# 実行
# -------------------------------------------
# D&D
$addDirStr = inputCondition
$Args | ForEach-Object {
    $item = Get-Item -LiteralPath $_
    if($item.PSIsContainer){
        # ファイルオブジェクトで渡すとパスが消える
        main $item.FullName $item.Name $addDirStr
    }
}