# =========================
# ファイル名一括置換
# =========================
$dirPath = "c:\temp\img"  # 置換対象ファイルがあるフォルダ

$nameFilePath = "c:\temp\name.txt" # ファイル名置換リスト
# ファイル名置換リスト
# aaaa,bbbbb
# cccc,ddddd
# dddd,eeeee

$enc = [System.Text.Encoding]::GetEncoding("UTF-8")

# ファイル名置換リスト読取
$namefile = [System.IO.File]::ReadAllText($nameFilePath, $enc)
[string[]]$nameAry = $namefile.split("`n")
$nameAry.Length

# 連想配列化
$hash = @{}
$namefile.split("`n") | ForEach-Object {
    $word = $_.split(",")
    $hash.add($word[0], $word[1].trim()) # 改行コードが含まれている？ので削除
}
$hash

# 置換対象取得
Get-ChildItem -Path $dirPath | ForEach-Object {
    if ($hash.ContainsKey($_.BaseName)) {
        Rename-Item -Path $_.FullName -NewName ($_.FullName).Replace($_.BaseName, $hash[$_.BaseName])
        # 置換後名は [regex]::Replace($_.FullName, $_.BaseName, $hash[$_.BaseName]) でもOK
    }
}