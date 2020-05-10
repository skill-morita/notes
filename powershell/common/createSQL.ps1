# ===========================================
# SQL作成
# ===========================================
# -------------------------------------------
# SQL文作成
# -------------------------------------------
<#
.SYNOPSIS 
    DELETE・INSERT文作成
.INPUTS
    table テーブル名
    hash key:フィールド名、val:値
    keyHash key:フィールド名、val:値
.OUTPUTS
    DELETE・INSERT文
#>
function createSqlDeleteInsert {
    param (
        [String]$table,
        [System.Collections.Hashtable]$hash,
        [System.Collections.Hashtable]$keyHash
    )
    [String]$sqlDel = createSqlDelete $table $keyHash
    [String]$sqlIns = createSqlInsert $table $hash
    return $sqlDel + $sqlIns
}

<#
.SYNOPSIS
    INSERT文作成
.INPUTS
    table テーブル名
    hash key:フィールド名、val:値
.OUTPUTS
    INSERT文
#>
function createSqlInsert {
    param (
        [String]$table,
        [System.Collections.Hashtable]$hash
    )
    $SQL_INSERT = "INSERT INTO {0} ({1}) VALUES ({2});"
    $key = $hash.Keys -join ","
    $val = $hash.Values -join ","
    return [String]::Format($SQL_INSERT, $table, $key, $val)
}

<#
.SYNOPSIS
    DELETE文作成
.INPUTS
    table テーブル名
    keyHash key:フィールド名、val:値
.OUTPUTS
    DELETE文
#>
function createSqlDelete {
    param (
        [String]$table,
        [System.Collections.Hashtable]$keyHash
    )
    $SQL_DELETE = "DELETE FROM {0} WHERE {1};"
    # WHERE句
    $whereStr = createWhere $keyHash
    return [String]::Format($SQL_DELETE, $table, $whereStr)
}

<#
.SYNOPSIS
    UPDATE文作成
.INPUTS
    table テーブル名
    hash key:フィールド名、val:値
    keyHash key:フィールド名、val:値
.OUTPUTS
    UPDATE文
#>
function createSqlUpdate {
    param (
        [String]$table,
        [System.Collections.Hashtable]$hash,
        [System.Collections.Hashtable]$keyHash
    )
    $SQL_UPDATE = "UPDATE {0} SET {1} WHERE {2};"
    # SET句
    $hash2 = $hash.GetEnumerator() | Select-Object { $_.Key + " = " + $_.Value } 
    $setStr = $hash2.Keys -join ","
    # WHERE句
    $whereStr = createWhere $keyHash
    return [String]::Format($SQL_UPDATE, $table, $setStr, $whereStr)
}

<#
.SYNOPSIS
    WHERE句作成
.INPUTS
    keyHash key:フィールド名、val:値
.OUTPUTS
    WHERE句
#>
function createWhere {
    param (
        [System.Collections.Hashtable]$keyHash
    )
    $keyHash2 = $keyHash.GetEnumerator() | Select-Object { $_.Key + " = " + $_.Value } 
    return $keyHash2.Values -join " AND "
}

# -------------------------------------------
# テスト値作成
# -------------------------------------------
function CreateTestVal {
    param (
        [string]$type,
        [int]$length
    )
    switch ($type) {
        { "character varying", "varchar", "character", "char" } { 
            # 文字列
            return CreateTestValString $type $length
        }
        { "smallint", "integer", "bigint", "decimal", "numeric"} { 
            # 数値
            return CreateTestValNumber $type $length
        }
        { "date", "timestamp" } { 
            # 日付
            return CreateTestValDate $type $length
        }
        Default { return $null }
    }
}

function CreateTestValString {
    param (
        [string]$type,
        [int]$length
    )
    $ary = @()
    # MAX桁
    $ary += "X" * $length
    $ary += "X" * ($length + 1)
    $ary += "あ" * $length
    $ary += "あ" * ($length + 1)
    # NULL、空文字、スペース
    $ary += ""
    $ary += $null
    $ary += " "
    # 記号
    $ary += ",' ./\=?!:; "
    $ary += "`""
    # カタカナ、半角カタカナ
    $ary += "ヲンヰヱヴーヾ・"
    $ary += "ｧｰｭｿﾏﾞﾟ"
    # 環境依存文字
    $ary += "㌶Ⅲ⑳㏾☎㈱髙﨑"
    # マッピングに差がある文字
    $ary += "¢£¬‖−〜―"
    # サロゲートペア
    $ary += "𠀋𡈽𡌛𡑮𡢽𠮟𡚴𡸴𣇄𣗄"
    # EUCのサーバで文字化けする文字
    $ary += "ソ能表"

    return $ary.GetEnumerator() | Select-Object { "'" + $_ + "'" }
}

function CreateTestValNumber {
    param (
        [string]$type,
        [int]$length
    )
    $ary = @()
    # 境界値
    switch ($type) {
        "smallint" { 
            $ary += -32769 # 下限値 + 1
            $ary += -32768 # 下限値
            $ary += 32767 # 上限値
            $ary += 32768 # 上限値 + 1
            break;
        }
        "integer" { 
            $ary += -2147483649 # 下限値 + 1
            $ary += -2147483648 # 下限値
            $ary += 2147483647 # 上限値
            $ary += 2147483648 # 上限値 + 1
            break;
        }
        "bigint" { 
            $ary += -9223372036854775809 # 下限値 + 1
            $ary += -9223372036854775808 # 下限値
            $ary += 9223372036854775807 # 上限値
            $ary += 9223372036854775808 # 上限値 + 1
            break;
        }
        "decimal" { 
            break;
        }
        "numeric" { 
            break;
        }
        Default { return $null }
    }
    # NULL、空文字、スペース
    $ary += ""
    $ary += $null
    $ary += " "
    return $ary
}

function CreateTestValDate {
    param (
        [string]$type,
        [int]$length
    )
    $ary = @()
    # TODO:書式
    return $ary
}