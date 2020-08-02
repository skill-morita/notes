# ===========================================
# フォルダ名にカテゴリ追加
# _カテゴリ_カテゴリ_元のフォルダ名
#
# 【使い方】
# .\renamedirAddCategory.ps1 -workpath C:\_git\__worktmp\______temp\
# ===========================================
# -------------------------------------------
# パラメータ
# -------------------------------------------
Param(
    [ValidateSet("C:\_git\__worktmp\" , "C:\_git\__worktmp\______temp\")]$workpath #作業フォルダ
)

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

function GetCategoryNameStringUser {
    param (
        $category_ary
    )

}

function GetCategoryAryUser {
    $category_ary = @()
    $category_ary += "一般"
    $category_ary += "要ログ"
    $category_ary += "フォロ限"
    $category_ary += "コミュ"
    return $category_ary
}

function FunctionName {
    param (
        $category_ary
    )
    $flg = "n";
    while ($flg -eq "y") {
        # 追加用の空配列
        $addAry = @()
        $addAry += GetCategoryName $category_ary
        $flg = (Read-Host "追加する？：y / n")
    }

    # カテゴリ連結
    return "_" + ($addAry -join "_")
}

# 選択肢を出力
function GetCategoryName {
    param (
        $category_ary
    )
    # otherの追加
    $category_ary += "other"

    # 入力時の説明文字列作成
    $category_explain = @()
    for ($i = 0; $i -lt $category_ary.Count; $i++) {
        $category_explain += $category_ary[$i] + "-" + $i 
    }
    $explain = "選択肢[" + ($category_explain -join " / " ) + "]"

    # 入力
    [int]$category_cd = (Read-Host $explain)
    [string]$category_nm = $category_ary[$category_cd]
    if ($category_nm -eq "other") {
        $category_nm = (Read-Host "名称を入力")
    }
    Write-Host $category_nm 
    return $category_nm 
}

function GetCategoryAryBackground {
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
    return $category_ary
}

# -------------------------------------------
# 実行
# -------------------------------------------
# パラメータ指定がない場合
if ($null -eq $workpath) {
    $workpath = "C:\_git\__worktmp\" 
}

$addDirStr = ""
[int]$category_ver = (Read-Host "1:配布者, 2:background")
switch ($category_ver) {
    1 { }
    2 { $addDirStr = GetCategoryName GetCategoryAryBackground }
    Default { }
}

$files = Get-TargetDirs $workpath $args
if ($null -eq $files) {
    exit
}
$files | ForEach-Object { RenameDirAddPrefix $_ $addDirStr }
Pause