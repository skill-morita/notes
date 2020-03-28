-- ===============================
-- 基本
-- ===============================
-- SQL*Plus起動
# sqlplus /nolog
-- ログ開始
SPOOL C:/temp/script.log
-- ～処理～
-- ログ終了
SPOOL OFF
-- SQL*Plus終了
exit

-- ===============================
-- まとめて
-- ===============================
-- SQL*Plus起動
# sqlplus /nolog > C:/temp/script.log
-- SQL*Plus終了
exit