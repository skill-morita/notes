-- ===============================
-- テーブル領域(EMExpressでも設定可能)
-- ===============================
----------------------------
-- # テーブル領域の作成
----------------------------
-- ファイルサイズ：4GB
-- 自動拡張 増分：100MB 最大ファイル・サイズ：無制限
-- エクステント管理：ローカル管理
-- 有効化
-- タイプ：永続
CREATE TABLESPACE TS_MAP 
DATAFILE 'D:\Oracle\19c\oradata\TS_MAP.dbf' SIZE 4G
AUTOEXTEND ON NEXT 100M MAXSIZE UNLIMITED 
EXTENT MANAGEMENT LOCAL 
ONLINE 
PERMANENT;

-- ファイルサイズ：200M
-- 自動拡張 増分：100MB 最大ファイル・サイズ：無制限
-- エクステント管理：ローカル管理
-- 有効化
-- タイプ：一時
CREATE TEMPORARY TABLESPACE TS_TEMP
TEMPFILE 'D:\Oracle\19c\oradata\TS_TEMP.dbf' SIZE 200M
AUTOEXTEND ON NEXT 50M MAXSIZE UNLIMITED
EXTENT MANAGEMENT LOCAL;

----------------------------
-- # テーブル領域の削除
----------------------------
-- INCLUDING CONTENTS：領域内の中身も削除
-- AND DATAFILES：物理ファイルも削除
-- CASCADE CONSTRAINTS：制約も削除
DROP TABLESPACE TS_MAP INCLUDING CONTENTS AND DATAFILES CASCADE CONSTRAINTS;

----------------------------
-- # テーブル領域の情報
----------------------------
SELECT
    TABLESPACE_NAME,
    STATUS,
    CONTENTS
FROM
    DBA_TABLESPACES
WHERE
    TABLESPACE_NAME LIKE 'TS_%';

-- STATUS：ONLINE(使用可能)、OFFLINE(使用不可)、READ ONLY(読取りのみ)
-- CONTENTS：UNDO、PERMANENT(永続)、TEMPORARY(一時)
-- その他もあり

----------------------------
-- # 読取り / 書込みの変更
----------------------------
-- 読取り/書込み
ALTER TABLESPACE TS_MAP READ WRITE;
-- 読取り
ALTER TABLESPACE TS_MAP READ ONLY;

----------------------------
-- # 使用可 / 不可の変更
----------------------------
ALTER TABLESPACE TS_MAP ONLINE;