-- ===============================
-- テスト中に使うSQL
-- ===============================
----------------------------
-- # INSERT/DELETE時にエラーを発生させる (インデックスを無効にする)
----------------------------
-- ## 状態確認
-- VALID インデックス有効
-- INVALID インデックス無効 
-- UNUSABLE インデックス使用禁止
SELECT
    index_name,
    STATUS
FROM
    user_indexes
WHERE
    index_name = 'IDX1';

-- ## 無効化
-- VALID -> UNUSABLE
-- テーブル移動
-- 既存のテーブル属性を継承したまま、同一表領域上の新しいセグメントで再構築
ALTER TABLE tbl1 MOVE;

-- ## 有効化
-- UNUSABLE -> VALID
-- インデックスを有効化する
ALTER INDEX idx1 REBUILD;

----------------------------
-- # INSERT/UPDATE/DELETE時にエラーを発生させる (例外を発生させる)
----------------------------
-- 例外を投げるトリガーを作成
CREATE OR REPLACE TRIGGER tst.tst_err
  BEFORE INSERT OR UPDATE OR DELETE ON tst.tbl1 FOR EACH ROW
BEGIN
  RAISE_APPLICATION_ERROR(-20000, 'エラー');
END;
/

-- トリガー削除
DROP TRIGGER tst.tst_err;

-- トリガー有効化/無効化
ALTER TRIGGER tst.tst_err ENABLE;
ALTER TRIGGER tst.tst_err DISABLE;

-- エラー確認
SHOW ERRORS
SHOW ERROR TRIGGER tst.tst_err

-- コンパイル
ALTER TRIGGER tst.tst_err COMPILE;