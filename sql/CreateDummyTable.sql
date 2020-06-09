-- ===============================
-- ダミーテーブルを作る
-- ===============================
----------------------------
-- Oracle : with句を使う
-- 動作確認：Oracle19c
----------------------------
with dummy1(cd, name) AS (
    SELECT 0, 'name1' FROM DUAL UNION ALL
    SELECT 1, 'name2' FROM DUAL UNION ALL
    SELECT 2, 'name3' FROM DUAL UNION ALL
    SELECT 3, 'name4' FROM DUAL UNION ALL
    SELECT 4, 'name5' FROM DUAL
), dummy2(cd, name) AS (
    SELECT 0, 'name1A' FROM DUAL UNION ALL
    SELECT 1, 'name2A' FROM DUAL UNION ALL
    SELECT 2, 'name3A' FROM DUAL UNION ALL
    SELECT 3, 'name4A' FROM DUAL UNION ALL
    SELECT 4, 'name5A' FROM DUAL
)
SELECT
    dummy1.cd,
    dummy1.name,
    dummy2.name
FROM dummy1
INNER JOIN  dummy2 ON dummy2.cd = dummy1.cd;
-- 「ORA-00928: SELECTキーワードがありません。」とエラーが出た時は、最後のSELECT文にも「UNION ALL」を追加している