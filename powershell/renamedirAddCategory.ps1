# ===========================================
# フォルダ名にカテゴリ追加
# _カテゴリ_カテゴリ_元のフォルダ名
# ===========================================
# -------------------------------------------
# 定数
# -------------------------------------------
$TARGET_DIR = "C:\_git\__worktmp\______temp"
# -------------------------------------------
# 関数
# -------------------------------------------
# 共通部品
. .\common\module.ps1

# フォルダ名頭に
function RenameDirAddPrefix {
    param (
        [System.IO.DirectoryInfo]$dirinfo,
        [string]$prefix
    )
    # フォルダ名変更
    $newnm = ""
    $oldnm = $dirinfo.Name
    if (($oldnm).Substring(0, 1) -eq "_") {
        $newnm = $prefix + $oldnm 
    } else {
        $newnm = $prefix + "_" + $oldnm 
    }

    Rename-Item -Path $dirinfo.FullName -NewName $newnm
    Write-ErrorLog $TARGET_DIR
}

# 追加カテゴリ名の入力
function GetCategoryNameString {
    # カテゴリ追加する場合↓----------------
    $category_ary = @()
    $category_ary += "室内"
    $category_ary += "室外"
    $category_ary += "抽象"
    $category_ary += "洋風"
    $category_ary += "和風"
    $category_ary += "中華"
    $category_ary += "現代"
    $category_ary += "植物"
    $category_ary += "シンプル"
    $category_ary += "かわいい"
    $category_ary += "きれい"
    $category_ary += "skydome"
    $category_ary += "other"

    $explain = "選択肢:"
    for ($i = 0; $i -lt $category_ary.Count; $i++) {
        $explain += ($category_ary[$i] + " " + $i + "/")
    }

    # 配布方法
    [int]$category_cd = (Read-Host $explain)
    [string]$category_nm = $category_ary[$category_cd]
    if ($category_nm -eq "other") {
        $category_nm = (Read-Host "名称を入力")
    }
    Write-Host $category_nm 
    # 追加用の空配列
    $addAry = @()
    $addAry += $category_nm
    # カテゴリ追加する場合↑----------------

    # カテゴリ連結
    return "_" + ($addAry -join "_")
}

# -------------------------------------------
# 実行
# -------------------------------------------
$addDirStr = GetCategoryNameString
$files = Get-TargetDirs $TARGET_DIR $args
if ($null -eq $files) {
    exit
}
$files | ForEach-Object { RenameDirAddPrefix $_ $addDirStr }
Pause