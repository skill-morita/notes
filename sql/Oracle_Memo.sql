-- ===============================
-- 色々雑多に
-- ===============================
----------------------------
-- # システム日付の取得 & フォーマット
----------------------------
SELECT
    TO_CHAR(SYSDATE, 'YYYYMMDD')
FROM
    DUAL;

----------------------------
-- # 文字列型のデータ長換算の確認
----------------------------
-- インスタンスごと
SELECT
    VALUE
FROM
    nls_database_parameters
WHERE
    parameter = 'NLS_LENGTH_SEMANTICS';

-- BYTE：バイト
-- CHAR：文字数

-- 列ごと
SELECT
    column_name,
    data_type, 
    data_length,
    char_used, 
    owner
FROM
    all_tab_columns
WHERE
    owner = 'testuser' -- ユーザー(≒スキーマ)
    AND column_name IN ('col1');  -- 列名

-- char_used が B：バイト
----------------------------
-- # 文字コードの確認
----------------------------
SELECT
    VALUE
FROM
    nls_database_parameters
WHERE
    parameter = 'NLS_CHARACTERSET';

-- JA16SJISTILDE：Shift-JIS


