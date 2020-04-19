# ===========================================
# フォルダ名にカテゴリ追加(ユーザー)
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

function main {
    # カテゴリ追加する場合↓----------------
    # 配布方法
    [int]$dist = (Read-Host 配布方法:1 一般、2 要ログ)
    [string]$distnm = ""
    switch ($dist) {
        1 {$distnm = "一般"; break}
        2 {$distnm = "要ログ"; break}
    }

    # ユーザー名
    [string]$usernm = (Read-Host 配布者)

    # 追加用の空配列
    $addAry = @()
    $addAry += $distnm
    $addAry += $usernm
    # カテゴリ追加する場合↑----------------

    # _配布方法_ユーザー名
    $addStr = "_" + ($addAry -join "_")

    # フォルダ一覧取得 再帰なし
    $dirList = Get-ChildItem -Path $TARGET_DIR -Directory

    $dirList | ForEach-Object {
        # フォルダ名変更
        $newnm = ""
        $oldnm = $_.Name
        if (($oldnm).Substring(0, 1) -eq "_") {
            $newnm = $addStr + $oldnm
        } else {
            $newnm = $addStr + "_" + $oldnm
        }
        Rename-Item -Path $_.FullName -NewName $newnm
        Write-ErrorLog
    }
}

# -------------------------------------------
# 実行
# -------------------------------------------
main