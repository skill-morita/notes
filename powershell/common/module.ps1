<#
.SYNOPSIS
    エラーログ出力
.PARAMETER logdirpath
    ログファイル出力フォルダ
.OUTPUTS
    デバッグログファイル
#>
function Write-ErrorLog {
    param (
        [string]$logdirpath
    )
    $msg = $Error[0].Exception.Message
    # 画面に表示
    # Write-Warning $msg
    # ログ出力
    Write-OutputLogfile $logdirpath "ERROR_" $msg

    $Error.Clear()
}

<#
.SYNOPSIS
    デバッグログ出力
.PARAMETER logdirpath
    ログファイル出力フォルダ
.PARAMETER msg
    出力メッセージ
.OUTPUTS
    デバッグログファイル
#>
function Write-DebugLog {
    param (
        [string]$logdirpath, 
        $msg
    )
    # 画面に表示
    Write-Debug $msg
    # ログ出力
    Write-OutputLogfile $logdirpath "DEBUG_" $msg
}

function Write-OutputLogfile {
    param (
        [string]$logdirpath, 
        [string]$prefix,
        $msg
    )
    # 実行中スクリプトファイル名
    $logPath = $logdirpath + "\" + $prefix + $script:MyInvocation.MyCommand.Name + ".log"
    $date = Get-DateString
    $logStr = ($date + $msg) 
    # ログ出力
    Write-Output $logStr | Out-File -FilePath $logPath -Force -Append
}

<#
.SYNOPSIS
    日付取得
#>
function Get-DateString {
    return Get-Date -format "[yyyy/MM/dd HH:mm:ss]"
}

<#
.SYNOPSIS
    ポーズ
#>
function Pause {
    if ($psISE) {
        $null = Read-Host 'Press Enter Key...'
    } else {
        Write-Host "Press Any Key..."
        (Get-Host).UI.RawUI.ReadKey()
    }
}

<#
.SYNOPSIS
    処理対象のフォルダオブジェクトを取得する。
    D&Dされたものか作業フォルダ配下の全フォルダ
.DESCRIPTION
    スクリプト実行ショートカットに-Fileのオプションを付けること。
    (スペース含むパスの対応)
    powershell.exe -File test.ps1
.PARAMETER targetDirpath
    作業フォルダ
.PARAMETER dropfiles
    D&Dで受け取った引数
.OUTPUTS
    System.IO.DirectoryInfo の配列
#>
function Get-TargetDirs {
    param (
        [string]$targetDirpath,
        [string[]]$dropDirs
    )
    $dirs = Get-DragDropFiles $dropDirs
    if ($dirs.Length -eq 0) {
        $explain = "配下の全フォルダを対象にします。OK:y NG:その他"
        [string]$confirm = (Read-Host "配下の")
        $dirs = Get-ChildItem -Path $targetDirpath -Directory 
    }
    return [System.IO.DirectoryInfo[]]$dirs
}

<#
.SYNOPSIS
    D&Dファイル
.DESCRIPTION
    スクリプト実行ショートカットに-Fileのオプションを付けること。
    (スペース含むパスの対応)
    powershell.exe -File test.ps1
.PARAMETER dropfiles
    D&Dで受け取った引数
.OUTPUTS
    System.IO.DirectoryInfo or System.IO.FileInfo の配列
#>
function Get-DragDropFiles {
    param (
        [string[]]$dropfiles
    )
    if ($dropfiles.Length -eq 0) {
        return $null
    } else {
        $fileObjs = $dropfiles | ForEach-Object { Get-Item -LiteralPath $_ }
        return $fileObjs
    }
}
