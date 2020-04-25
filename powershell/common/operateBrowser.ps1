# ===========================================
# ブラウザの画面操作をする
# 
# [PowershellでInternetExplorerを操作する \- Qiita](https://qiita.com/flasksrw/items/a1ff5fbbc3b660e01d96)
# [MSHTML Reference - Internet Explorer C\+\+ Reference \(Windows\) \| Microsoft Docs](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/hh801968(v=vs.85))
# ===========================================
# -------------------------------------------
# 定数
# -------------------------------------------

# -------------------------------------------
# IE操作
# -------------------------------------------
# 既に開いているIEを取得
function GetExistIE {
    # シェルを取得
    $shell = New-Object -ComObject Shell.Application

    # IE取得
    $ie = $shell.Windows() |
        Where-Object { $_.Name -eq "Internet Explorer" } |
            Where-Object { $_.LocationURL -like $ACCESS_URL } |
                Select-Object -First 1

    return $ie
}

# IEウインドウをアクティブ化する
function ActivateIE {
    param(
        $ie
    )
    Add-Type -AssemblyName Microsoft.VisualBasic
    $window_process = Get-Process -Name "iexplore" | Where-Object { $_.MainWindowHandle -eq $ie.HWND }
    [Microsoft.VisualBasic.Interaction]::AppActivate($window_process.ID) | Out-Null
}

# Frame内DOMの取得
function GetFrameDocument {
    [CmdletBinding()]
    param (
        $ie, $frameName
    )
    # DOM取得 
    $document = OverrideHTMLDocument($ie.Document)
    # 入力フレーム取得
    $frame = $document.getElementsByName($frameName)
    # frame内のDOM
    return OverrideHTMLDocument $frame.contentDocument
}

# DOMのメソッドを再定義
# [PowershellでInternetExplorerを操作する \- Qiita](https://qiita.com/flasksrw/items/a1ff5fbbc3b660e01d96)
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
# TEXT、SELECT入力
function SetElm {
    param (
        [mshtml.IHTMLElement]$elm,
        $val
    )
    $elm.value = [string]$val
    $elm.fireEvent("onchange")
    Start-Sleep 1
}

# TEXT、SELECT初期化
function ClearElm {
    param (
        [mshtml.IHTMLElement]$elm
    )
    $elm.value = ""
    $elm.fireEvent("onchange")
    Start-Sleep 1
}

# 最大値入力
function CreateMaxLength {
    param (
        [mshtml.IHTMLElement]$elm
    )
    return "X" * $elm.maxLength
}
