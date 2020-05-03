# ログ出力
function Write-ErrorLog {
    param (
        $logdirpath
    )
    # 実行中スクリプトファイル名
    $logPath = $logdirpath + "ERROR_" + $script:MyInvocation.MyCommand.Name + ".log"
    $logStr = $Error[0].Exception.Message
    Write-Output $logStr | Out-File -FilePath $logPath -Force -Append
    $Error.Clear()
}

function Write-DebugLog {
    param (
        $logdirpath, $msg
    )
    # 実行中スクリプトファイル名
    $logPath = $logdirpath + "DEBUG_" + $script:MyInvocation.MyCommand.Name + ".log"
    $logStr = $msg
    Write-Output $logStr | Out-File -FilePath $logPath -Force -Append
}

function Get-DateString {
    return Get-Date -format "yyyy/MM/dd HH:mm:ss"
}