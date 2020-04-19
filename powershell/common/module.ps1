# ログ出力
function ErrorLog {
    # 実行中スクリプトファイル名
    $logPath = $TARGET_DIR + $script:MyInvocation.MyCommand.Name + ".log"
    Write-Output $Error[0].Exception.Message | Out-File -FilePath $logPath -Force
    $Error.Clear()
}
