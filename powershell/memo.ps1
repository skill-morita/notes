#======================================
# 基本
#======================================
# コンソール表示
Write-Host "Hello World!"

# 変数は変数だけ書くと表示される
# スネークケース
$usernm = "Mike"
$usernm

# 型指定
[int]$age = 1
$age

# 文字列連結などは括弧で括らないと+がパラメータ扱いでエラーになる
Write-Host ("Hello" + "World!")

# 関数の書き方
function func {
    param (
        $usernm,
        $age
    )
    Write-Host ($usernm + "は" + $age + "歳です")
} 
function func_simple ($usernm, $age) {
    Write-Host ($usernm + "は" + $age + "歳です")
} 
# 引数を括弧でくくったり、カンマで区切ると配列の扱いになったリするので注意。
# あくまでコマンド
func $usernm $age 

# ループ処理
# [Powershellのforeach\(ForEach\-Object\)の使い方 \| マイクロソフ党ブログ](https://microsoftou.com/ps-foreach/#toc2)
# [PowerShellの使い方\(オブジェクト操作編\) \- Qiita](https://qiita.com/Kirito1617/items/bd3937fb26c668eca078)
# [PowerShell の Sort\-Object Tips](https://www.vwnet.jp/Windows/PowerShell/2017032901/Sort-Object_Tips.htm)
function main_loop {
    $ary = @("apple", "orange")

    # ForEach-Object 基本
    # $_はパイプで受け取ったオブジェクトを参照する変数
    $ary | ForEach-Object { Write-Host $_ }

    # ForEach-Object の書き換え
    $ary | % { Write-Host $_ }

    # foreach 大量のデータを高速に処理したい場合によい
    foreach ($item in $ary) {
        Write-Host $item
    }

    # ForEach、Where、Select、Sortの使い分けはLinqを思い出す
    # ForEach：返り値なしの処理
    # Where：絞り込み
    $ary | Where-Object { $_ -like "a*" }
    # Select：加工(オブジェクトから１つのプロパティにするとか)
    $ary | Select-Object { $_ + "_add" }
    # Sort：並び替え
    $ary | Sort-Object -Descending

    # ファイル内をGrepする
    Select-String -Path "C:/temp/file.txt" -Pattern "/d{4}" # 指定ファイル
    Select-String "C:/temp/file.txt" -Pattern "/d{4}" # 指定ファイル
    Select-String -Path "C:/temp/*" -Pattern "/d{4}" # 指定フォルダ配下のファイル
    Select-String -Path "C:/temp/*.txt" -Pattern "/d{4}" # 指定フォルダ配下のtxtファイル
}

#======================================
# 配列・リスト
#======================================
# 配列
$ary = @("apple", "orange")
$ary

# 動的配列
$ary2 = @()
$ary2 += "apple"
$ary2 += "orange"

# 連想配列
$hash = @{ }
$hash.add("key", "val")
$hash.Remove("key")
$hash2 = @{
    "key1" = "val1"
    "key2" = "val2"
}
$hash2["key1"] = "val1A"
$hash2.Keys
$hash2.ContainsKey("key1")
$hash2.Values
$hash2.ContainsValue("key1")

# 簡単にリストを定義する
$CountryList = @"
USA
Germany
China
India
Japan
"@ -Split "`r`n"
$CountryList

#======================================
# ログなど
#======================================
function main_measyre ($path) {
    # 時間を計測する
    Measure-Command {
        Get-ChildItem $path\*.xls*
    }
    # 数を数える
    (Get-ChildItem -Path $path\*.xls* -File | Measure-Object).Count
}

# ログ出力
function Write-ErrorLog {
    # 実行中スクリプトファイル名
    $logPath = $TARGET_DIR + $script:MyInvocation.MyCommand.Name + ".log"
    Write-Output $Error[0].Exception.Message | Out-File -FilePath $logPath -Force
    $Error.Clear()
}

#======================================
# 検索・置換・絞り込み
#======================================
function main_search {
    # 正規表現検索
    $baseName = "PG1001_メインメニュー"
    $pgcd = [regex]::Matches($baseName, "^PG[0-9]*")

    # 置換
    $newPgcd = $pgcd -replace "PG", "Program"
    $newPgcd

    # 指定文字列を含むものを除外する
    $dirPath = "C:\TEMP"
    $file = Get-ChildItem $dirPath\*.xml* | Where-Object { $_ -NotLike "aaa*" } # 文字列のダブルクオテーションと{}を忘れない
    # ファイルの中身を検索
    $grepFile = $file | Select-String "bbb" -Encoding default -CaseSensitive
    $grepFile
}
#======================================
# ファイル操作
#======================================
function main_file ($path) {
    # 指定フォルダ配下のファイルオブジェクトリスト取得
    $files = Get-ChildItem $path\*.xls*

    # 拡張子を除くファイル名取得
    $baseNames = $files.BaseName
    $baseNames

    # フォルダ作成
    $dirPath = "C:\TEMP"
    New-Item $path -ItemType Directory -Force

    # 存在確認しつつフォルダ作成
    $dirPath = "C:\TEMP"
    if (!(Test-Path $dirPath)) {
        New-Item $dirPath -ItemType Directory
    }

    # 配下のファイル取得
    $dirPath = "C:\TEMP"
    $file = Get-ChildItem $dirPath\*.xml*
    $file # 表示

    # 絞り込みながら再帰でサブフォルダ内も検索
    $dirPath = "C:\TEMP"
    $file = Get-ChildItem $dirPath\*.xml* -Recurse -Filter aaa* `
        $file | Select-Object FullName # ファイルパス表示
    $file | Get-Content # ファイル読取
}

#======================================
# パス操作
#======================================
function main_path ($files) {
    # 親フォルダのパスを取得したい
    $files | ForEach-Object { Split-Path $_.FullName -Parent }
}

#======================================
# XML
#======================================
function main_xml {
    # XML取得
    $path = "C:\TEMP\a.xml"
    $xmlDoc = [xml](Get-Content $path -enc UTF8)
    $xmlNav = $xmlDoc.CreateNavigator()

    # XPathで取得(「AAAA」が含まれているliを持つdiv)
    $xmlNav.Select(".//div[contains(./ol/li/text(), 'AAAA')]") | ForEach-Object {
        # 属性取得(divの名前)
        $_.GetAttribute("Name", "")
        # 値取得(liの中身)
        $_.Select("./ol/li/text()").Value
    }
}

#======================================
# 出入力
#======================================
function main_inout() {
    # 入力したい
    $usernm = (Read-Host ユーザー名の入力)
    $usernm
}

# ps1ファイルにD & D
# $args(=コマンドライン引数)をfuction内に書くとローカル変数扱いになって取得できない？
$args | ForEach-Object { # ArgsがD&Dしたファイルやフォルダ
    Write-Host $_.GetType()
    $item = Get-Item -LiteralPath $_ # パスにスペースが有っても1まとまりとして認識する
    Write-Host $item.FullName
}
