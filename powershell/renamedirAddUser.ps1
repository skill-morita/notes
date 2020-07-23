# ===========================================
# フォルダ名にカテゴリ追加(ユーザー)
# _配布方法_ユーザー名_元のフォルダ名
#
# 【使い方】
# .\renamedirAddUser.ps1 -workpath C:\_git\__worktmp\______temp\
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

    # 配布方法 の入力
    $categoryAry = GetCategoryAryUser
    [string]$distnm = GetCategoryName $categoryAry

    # ユーザー名 の入力
    [string]$usernm = (Read-Host 配布者)

    # _配布方法_ユーザー名 に整形
    $addAry = @()
    $addAry += $distnm
    $addAry += $usernm
    $addStr = "_" + ($addAry -join "_")

    # フォルダのリネーム
    RenameDirList $workpath_ $addStr
}

# フォルダのリネーム
function RenameDirList {
    param (
        $workpath_,
        $addStr_
    )
    # フォルダ一覧取得 再帰なし
    $dirList = Get-ChildItem -Path $workpath_ -Directory

    $dirList | ForEach-Object {
        $oldnm = $_.Name
        # 先頭文字列にハイフンがなければ追加
        $newnm = ""
        if (($oldnm).Substring(0, 1) -eq "_") {
            $newnm = $addStr_ + $oldnm
        } else {
            $newnm = $addStr_ + "_" + $oldnm
        }
        # フォルダ名変更
        Rename-Item -Path $_.FullName -NewName $newnm
        Write-ErrorLog $workpath_
    }
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

# 配布方法選択肢
function GetCategoryAryUser {
    $category_ary = @()
    $category_ary += "一般"
    $category_ary += "要ログ"
    $category_ary += "フォロ限"
    $category_ary += "コミュ"
    return $category_ary
}

# -------------------------------------------
# 実行
# -------------------------------------------
# パラメータ指定がない場合
if ($null -eq $workpath) {
    $workpath = "C:\_git\__worktmp\" 
}
main $workpath