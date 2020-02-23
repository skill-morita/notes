# コンソール表示
Write-Host "Hello World!"

# 変数は変数だけ書くと表示される
$files

# 指定フォルダ配下のファイルオブジェクトリスト取得
$path = "C:\TEMP"
$files = Get-ChildItem $path\*.xls*

# 拡張子を除くファイル名取得
$baseNames = $files.BaseName

# 正規表現検索
$baseName = "PG1001_メインメニュー"
$pgcd = [regex]::Matches($baseName, "^PG[0-9]*")

# Foreach($_はパイプで受け取ったオブジェクトを参照する変数)
$pgcds = $baseNames.foreach{
    [regex]::Matches($_, "^PG[0-9]*")
}

# 置換
$newPgcd = $pgcd -replace "PG", "Program"

# フォルダ作成
$dirPath = "C:\TEMP"
New-Item $path -ItemType Directory -Force

# 存在確認しつつフォルダ作成
$dirPath = "C:\TEMP"
if(!(Test-Path $dirPath)){
    New-Item $dirPath -ItemType Directory
}

# XML取得
$path = "C:\TEMP\a.xml"
$xmlDoc = [xml](Get-Content $path -enc UTF8)
$xmlNav = $xmlDoc.CreateNavigator()

# XPathで取得(「AAAA」が含まれているliを持つdiv)
$xmlNav.Select(".//div[contains(./ol/li/text(), 'AAAA')]") | ForEach-Object{
    # 属性取得(divの名前)
    $_.GetAttribute("Name", "")
    # 値取得(liの中身)
    $_.Select("./ol/li/text()").Value
}

# 簡単にリストを定義する
$CountryList = @"
USA
Germany
China
India
Japan
"@ -Split "`r`n"

# 時間を計測する
Measure-Command {
    Get-ChildItem $path\*.xls*
    New-Item $path\$newPgcd -ItemType Directory -Force
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

# 指定文字列を含むものを除外する
$dirPath = "C:\TEMP"
$file = Get-ChildItem $dirPath\*.xml* | Where-Object {$_ -NotLike "aaa*"} # 文字列のダブルクオテーションと{}を忘れない
# 中身を検索
$grepFile = $file | Select-String "bbb" -Encoding default -CaseSensitive
$grepFile