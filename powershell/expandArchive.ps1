﻿$TARGET_DIR = "C:\_git\__worktmp\"
$EXT = ".zip"

function main {
    # 圧縮ファイル一覧取得 再帰なし
    $zipfile = Get-ChildItem -Path $TARGET_DIR -File -Filter *$EXT
    # $zipfile
    $zipfile | ForEach-Object {
        $dirnm = ($_.FullName).Replace($EXT, "")
        # zip展開
        Expand-Archive -Path $_.FullName -DestinationPath $dirnm 
        ErrorLog
        # 展開したフォルダにzipを移動
        Move-Item $_.FullName $dirnm
    }
}

# ログ出力
function ErrorLog {
    # 実行中スクリプトファイル名
    $logPath = $TARGET_DIR + $script:MyInvocation.MyCommand.Name + ".log"
    Write-Output $Error[0].Exception.Message | Out-File -FilePath $logPath -Force
    $Error.Clear()
}

# ===========================================
# 実行
# ===========================================
main