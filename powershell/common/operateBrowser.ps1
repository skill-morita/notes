# ===========================================
# ブラウザの画面操作をする
# ===========================================
# -------------------------------------------
# IE操作
# -------------------------------------------
<#
.SYNOPSIS
    IEを起動して取得
.INPUTS 
    access_url URL
.OUTPUTS
    IEオブジェクト
.NOTES
    [PowershellでInternetExplorerを操作する \- Qiita](https://qiita.com/flasksrw/items/a1ff5fbbc3b660e01d96)
#>
function CreateIE {
    [CmdletBinding()]
    param (
        [string]$access_url
    )
    $ie = New-Object -ComObject InternetExplorer.Application
    $ie.Visible = $true
    $ie.Navigate($access_url, 4)

    $shell = New-Object -ComObject Shell.Application
    while ($ie.Document -isnot [mshtml.HTMLDocumentClass]) {
        $ie = $shell.Windows() | Where-Object { $_.HWND -eq $hwnd }
    }
    return $ie
}

<#
.SYNOPSIS
    既に開いているIEを取得
.INPUTS 
    access_url URL
.OUTPUTS
    IEオブジェクト
.NOTES
    [PowershellでInternetExplorerを操作する \- Qiita](https://qiita.com/flasksrw/items/a1ff5fbbc3b660e01d96)
#>
function GetExistIE {
    [CmdletBinding()]
    param(
        [string]$access_url
    )
    # シェルを取得
    $shell = New-Object -ComObject Shell.Application
    # IE取得
    $ie = $shell.Windows() |
        Where-Object { $_.Name -eq "Internet Explorer" } |
            Where-Object { $_.LocationURL -like $access_url } |
                Select-Object -First 1

    return $ie
}

<#
.SYNOPSIS
    IEウインドウを終了する
.INPUTS 
    ie IEオブジェクト
#>
function CloseIE {
    [CmdletBinding()]
    param(
        $ie
    )
    $ie.quit()
}

<#
.SYNOPSIS
    IEウインドウをアクティブ化する
.INPUTS 
    ie IEオブジェクト
.NOTES
    [PowershellでInternetExplorerを操作する \- Qiita](https://qiita.com/flasksrw/items/a1ff5fbbc3b660e01d96)
#>
function ActivateIE {
    [CmdletBinding()]
    param(
        $ie
    )
    Add-Type -AssemblyName Microsoft.VisualBasic
    $window_process = Get-Process -Name "iexplore" | Where-Object { $_.MainWindowHandle -eq $ie.HWND }
    [Microsoft.VisualBasic.Interaction]::AppActivate($window_process.ID) | Out-Null
}

<#
.SYNOPSIS
    IEウインドウを読込待機
.INPUTS 
    ie IEオブジェクト
.NOTES
    [PowerShellからIEを操作 \- Qiita](https://qiita.com/fujimohige/items/5aafe5604a943f74f6f0)
#>
function WaitIE {
    [CmdletBinding()]
    param (
        $ie
    )
    While ($ie.Busy)
    { Start-Sleep -s 1 } 
}

<#
.SYNOPSIS
    IEウインドウを再読込
.INPUTS 
    ie IEオブジェクト
#>
function ReloadIE {
    [CmdletBinding()]
    param (
        $ie
    )
    $ie.Refresh()
}

<#
.SYNOPSIS
    Frame内DOMの取得
.INPUTS 
    ie IEオブジェクト
    frameName FRAME名
.OUTPUTS
    DOMオブジェクト
#>
function GetFrameDocument {
    [CmdletBinding()]
    param (
        $ie, 
        [string]$frameName
    )
    # DOM取得 
    $document = OverrideHTMLDocument($ie.Document)
    # 入力フレーム取得
    $frame = $document.getElementsByName($frameName)
    # frame内のDOM
    return OverrideHTMLDocument $frame.contentDocument
}

<#
.SYNOPSIS
    DOMのメソッドを再定義
.INPUTS 
    document mshtml.HTMLDocumentClass
.OUTPUTS
    mshtml.HTMLDocumentClass
.NOTES
    [PowershellでInternetExplorerを操作する \- Qiita](https://qiita.com/flasksrw/items/a1ff5fbbc3b660e01d96)
#>
function OverrideHTMLDocument () {
    [CmdletBinding()]
    param (
        [mshtml.HTMLDocumentClass]$document
    )
    #----------------------------------------------------------
    # getElement
    $doc = $document | Add-Member -MemberType ScriptMethod -Name "getElementById" -Value {
        param($Id)
        [System.__ComObject].InvokeMember(
            "getElementById", [System.Reflection.BindingFlags]::InvokeMethod, $null, $this, $Id
        ) | Where-Object { $_ -ne [System.DBNull]::Value }
    } -Force -PassThru

    $doc | Add-Member -MemberType ScriptMethod -Name "getElementsByClassName" -Value {
        param($ClassName)
        [System.__ComObject].InvokeMember( 
            "getElementsByClassName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $this, $ClassName
        ) | Where-Object { $_ -ne [System.DBNull]::Value }
    } -Force

    $doc | Add-Member -MemberType ScriptMethod -Name "getElementsByTagName" -Value {
        param($TagName)
        [System.__ComObject].InvokeMember( "getElementsByTagName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $this, $TagName
        ) | Where-Object { $_ -ne [System.DBNull]::Value }
    } -Force

    $doc | Add-Member -MemberType ScriptMethod -Name "getElementsByName" -Value {
        param($Name)
        [System.__ComObject].InvokeMember(
            "getElementsByName", [System.Reflection.BindingFlags]::InvokeMethod, $null, $this, $Name
        ) | Where-Object { $_ -ne [System.DBNull]::Value }
    } -Force

    #----------------------------------------------------------
    # button
    $doc | Add-Member -MemberType ScriptMethod -Name "click" -Value {
        [System.__ComObject].InvokeMember(
            "click",
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $this,
            $null
        ) | Where-Object { $_ -ne [System.DBNull]::Value }
    } -Force

    return $doc
}

# -------------------------------------------
# フォーム入力
# -------------------------------------------
<#
.SYNOPSIS
    TEXT,TEXTAREAに値をセットする
.INPUTS 
    elm エレメント
    val 値
#>
function SetElm {
    [CmdletBinding()]
    param (
        [mshtml.IHTMLElement]$elm,
        $val
    )
    if ($elm -eq $null) {
        Write-Host "null"
        break;
    }

    if ($elm.tagName -eq "textarea") {
        $elm.textContent = [string]$val
    } else {
        $elm.value = [string]$val
    }
    if ($elm.dispatchEvent) {
        # TODO: IE9以降
    } elseif ($elm.fireEvent) {
        # IE8以前
        $elm.fireEvent("onchange")
    }
    Start-Sleep 1
}

<#
.SYNOPSIS
    RADIOに値をセットする
.INPUTS 
    elm エレメント配列
    val インデックス
#>
function SetRadio {
    [CmdletBinding()]
    param (
        [System.Object[]]$elms,
        [int]$cnt
    )
    $elm = $elms[$cnt]
    $elm.checked = $true
    if ($elm.dispatchEvent) {
        # TODO: IE9以降
    } elseif ($elm.fireEvent) {
        # IE8以前
        $elm.fireEvent("onchange")
    }
    Start-Sleep 1
}

<#
.SYNOPSIS
    TEXT,TEXTAREAを初期化
.INPUTS 
    elm エレメント
#>
function ClearElm {
    param (
        [mshtml.IHTMLElement]$elm
    )
    SetElm $elm ""
}

<#
.SYNOPSIS
    最大文字列取得
.INPUTS 
    elm エレメント
    val 繰り返し文字列
.OUTPUTS
    最大文字列
#>
function CreateMaxLength {
    [CmdletBinding()]
    param (
        [mshtml.IHTMLElement]$elm, 
        $val
    )
    return $val * $elm.maxLength
}
