# ログ出力
function Write-ErrorLog {
    # 実行中スクリプトファイル名
    $logPath = $TARGET_DIR + "ERROR_" + $script:MyInvocation.MyCommand.Name + ".log"
    Write-Output $Error[0].Exception.Message | Out-File -FilePath $logPath -Force -Append
    $Error.Clear()
}

function Write-DebugLog ($msg) {
    # 実行中スクリプトファイル名
    $logPath = $TARGET_DIR + "DEBUG_" + $script:MyInvocation.MyCommand.Name + ".log"
    Write-Output $msg | Out-File -FilePath $logPath -Force -Append
}
