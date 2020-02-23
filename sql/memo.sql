----------------------------
--列名取得
----------------------------
--PostgreSQL
SELECT * 
FROM information_schema.columns
WHERE table_catalog = 'TestDB' --DB名
AND table_schema = 'public' --スキーマ名
AND data_type = 'character varying' --型名
AND column_name LIKE '%_ID%'; --検索したいカラム名

----------------------------
--同じ列名のデータを検索
----------------------------
--PostgreSQL
SELECT * 
FROM (
     SELECT aaa AS ColA, 'TblA' FROM TblA
     UNION SELECT aaa2 AS ColA, 'TblB' FROM TblB
     UNION SELECT aaa3 AS ColA, 'TblC' FROM TblC
) Tbl1
WHERE Tbl1.ColA LIKE '%flg%'