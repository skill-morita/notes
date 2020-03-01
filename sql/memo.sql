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
--¥d tblAでも可能

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

----------------------------
--既存データをSELECT⇒INSERTをしたい
----------------------------
--PostgreSQL
INSERT INTO tblA(colA, colB)
SELECT genelate_series(-100, -10) AS colA
, '0001' AS colB;

----------------------------
--UPSERT
----------------------------
--PostgreSQL 9.5より前
WITH upsert AS (
	UPDATE TblA SET ColA = val_ColA
	WHERE ColKey = val_Key
	RETURNING val_Key
)
INSERT INTO TblA (ColKey, ColA)
	SELECT val_Key, val_ColA From TblA
	WHERE NOT EXISTS (
		SELECT ColKey FROM upsert
);

----------------------------
--UPDATEでTableAの値をTableBに反映する
----------------------------
--PostgreSQL
UPDATE TableA
SET
	colA = TableB.colB,
	colA2 = TableB.colB2
FROM TableB
WHERE TableA.id = TableB.id;
--MySql
UPDATE TableA, TableB
SET
	TableA.colA = TableB.colB,
	TableA.colA2 = TableB.colB2
FROM TableB
WHERE TableA.id = TableB.id;