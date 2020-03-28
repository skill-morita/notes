# sqlplus / as sysdba --ログイン
@?/rbms/admin/execemx emx --Flash版に変更

# sqlplus / as sysdba --ログイン
@?/rbms/admin/execemx omx --JET版に変更

-- > セッションが変更されました。
-- > レコードが選択されませんでした。
-- > 旧 1:select nvl('&1','omx') p1 from dual
-- > 新 1:select nvl('emx','omx') p1 from dual
-- > P1
-- > -----
-- > emx
-- > PL/SQLプロシージャが正常に完了しました。
-- > セッションが変更されました。