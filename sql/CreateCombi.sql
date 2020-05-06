-- ===============================
-- 組み合わせを求める
-- 動作確認：PostgreSQL9.3
-- ===============================
----------------------------
-- データ作成
----------------------------
INSERT INTO pattern (pattern_id) VALUES (1);
INSERT INTO pattern (pattern_id) VALUES (2);
INSERT INTO pattern (pattern_id) VALUES (3);
INSERT INTO sel_youbi (pattern_id, youbi) VALUES (1, '日');
INSERT INTO sel_youbi (pattern_id, youbi) VALUES (1, '月');
INSERT INTO sel_youbi (pattern_id, youbi) VALUES (2, '日');
INSERT INTO sel_youbi (pattern_id, youbi) VALUES (3, '月');
INSERT INTO sel_youbi (pattern_id, youbi) VALUES (3, '火');
INSERT INTO sel_time (pattern_id, timezone) VALUES (1, '午前');
INSERT INTO sel_time (pattern_id, timezone) VALUES (2, '午前');
INSERT INTO sel_time (pattern_id, timezone) VALUES (3, '午前');
INSERT INTO sel_time (pattern_id, timezone) VALUES (3, '午後');
INSERT INTO sel_place (pattern_id, place) VALUES (1, '会社');
INSERT INTO sel_place (pattern_id, place) VALUES (1, 'リビング');
INSERT INTO sel_place (pattern_id, place) VALUES (2, 'カフェ');
INSERT INTO sel_place (pattern_id, place) VALUES (2, 'リビング');

----------------------------
-- 3要素すべてを使った組み合わせ
----------------------------
--パターンID：1の全組み合わせを求める。
SELECT
    pattern_id,
    youbi,
    timezone,
    place
FROM
    pattern
    INNER JOIN sel_youbi ON pattern.pattern_id = sel_youbi.pattern_id
    INNER JOIN sel_time ON pattern.pattern_id = sel_time.pattern_id
    INNER JOIN sel_place ON pattern.pattern_id = sel_place.pattern_id
WHERE
    pattern.pattern_id = 1;

----------------------------
-- 一部の要素の一部の値を指定した組み合わせ
----------------------------
--[曜日, 時間帯, 場所] = [月 or 火, 午後, 指定なし]の組み合わせを持つパターンIDを取得する
SELECT
    pattern_id,
    youbi,
    timezone
FROM
    pattern
    INNER JOIN sel_youbi ON pattern.pattern_id = sel_youbi.pattern_id
    INNER JOIN sel_time ON pattern.pattern_id = sel_time.pattern_id
    LEFT OUTER JOIN sel_place ON pattern.pattern_id = sel_place.pattern_id
WHERE
    sel_youbi.youbi IN ('月', '火')
    AND sel_time.timezone = '午後'
    AND sel_place.place IS NULL;