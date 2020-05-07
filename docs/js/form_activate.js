/** ======================================================
 * フォームの活性非活性
 * ====================================================== */
// -------------------------------------------------------
// 関数
// -------------------------------------------------------
/**
 * 非表示
 * @param {boolean} on true:非表示
 * @description ブロックがなくなる
 */
function setNonDisplayForms(on) {
    // テキストボックス
    setNonDisplay(document.getElementById("text1"), on);
    // ラジオボタン
    var radioElms = document.getElementsByName("radio1");
    for (let i = 0; i < radioElms.length; i++) {
        setNonDisplay(radioElms[i], on);
    }
    // チェックボックス
    var checkElms = document.getElementsByName("check1");
    for (let i = 0; i < checkElms.length; i++) {
        setNonDisplay(checkElms[i], on);
    }
    // セレクト
    setNonDisplay(document.getElementById("select1"), on);
}

/**
 * 非表示
 * @param {boolean} on true:非表示
 * @description ブロックは残る
 */
function setNonVisibilityForms(on) {
    // テキストボックス
    setNonVisibility(document.getElementById("text1"), on);
    // ラジオボタン
    var radioElms = document.getElementsByName("radio1");
    for (let i = 0; i < radioElms.length; i++) {
        setNonVisibility(radioElms[i], on);
    }
    // チェックボックス
    var checkElms = document.getElementsByName("check1");
    for (let i = 0; i < checkElms.length; i++) {
        setNonVisibility(checkElms[i], on);
    }
    // セレクト
    setNonVisibility(document.getElementById("select1"), on);
}

/**
 * 無効化
 * @param {boolean} on true:無効化
 * @description 入力 ：不可、submit:値が送信されない
 */
function setDisabledForms(on) {
    // テキストボックス
    setDisabled(document.getElementById("text1"), on);
    // ラジオボタン
    var radioElms = document.getElementsByName("radio1");
    for (let i = 0; i < radioElms.length; i++) {
        setDisabled(radioElms[i], on);
    }
    // チェックボックス
    var checkElms = document.getElementsByName("check1");
    for (let i = 0; i < checkElms.length; i++) {
        setDisabled(checkElms[i], on);
    }
    // セレクト
    // @ts-ignore
    setDisabled(document.getElementById("select1"), on);
}

/**
 * 読取専用化
 * @param {boolean} on true:読取専用化
 * @description 入力 ：不可、submit:値が送信される
 */
function setReadonlyForms(on) {
    // テキストボックス
    // @ts-ignore
    setReadonlyText(document.getElementById("text1"), on);
    // ラジオボタン
    setReadonlyCheckOrRadio(document.getElementsByName("radio1"), on);
    // チェックボックス
    setReadonlyCheckOrRadio(document.getElementsByName("check1"), on);
    // セレクト
    setReadonlySelect(document.getElementById("select1"), on);
}

/**
 * 非表示
 * @param {HTMLElement} elm
 * @param {boolean} on true:非表示
 * @description ブロックがなくなる
 */
function setNonDisplay(elm, on) {
    if (on) {
        elm.style.display = 'none'; //非表示
    } else {
        elm.style.display = ''; //表示
        // blockだと改行されてしまう？
    }
}

/**
 * 非表示
 * @param {HTMLElement} elm
 * @param {boolean} on true:非表示
 * @description ブロックは残る
 */
function setNonVisibility(elm, on) {
    if (on) {
        elm.style.visibility = 'hidden'; //非表示
    } else {
        elm.style.visibility = 'visible'; //表示
    }
}

/**
 * 無効化
 * @param {HTMLInputElement} elm
 * @param {boolean} on true:無効化
 * @description 入力 ：不可、submit:値が送信されない
 */
function setDisabled(elm, on) {
    if (on) {
        elm.disabled = true;
    } else {
        elm.disabled = false;
    }
}

/**
 * 読取専用化
 * @param {HTMLInputElement} elm テキストボックス
 * @param {boolean} on true:読取専用化
 * @description 入力 ：不可、submit:値が送信される
 */
function setReadonlyText(elm, on) {
    if (on) {
        elm.readOnly = true;
        elm.style.backgroundColor = '#CCCCCC'; //グレー
    } else {
        elm.readOnly = false;
        elm.style.backgroundColor = '#FFFFFF'; //白
    }
}

/**
 * 読取専用化
 * @param {HTMLCollection} elms Name指定したチェックボックスorラジオボタン
 * @param {boolean} on true:読取専用化
 * @description 入力 ：不可、submit:値が送信される
 */
function setReadonlyCheckOrRadio(elms, on) {
    for (let i = 0; i < elms.length; i++) {
        let item = elms[i];
        if (on && !item.checked) {
            item.disabled = true;
        } else {
            item.disabled = false;
        }
    }
}

/**
 * 読取専用化
 * @param {HTMLSelectElement} elm SELECTエレメント
 * @param {boolean} on true:読取専用化
 * @description 入力 ：不可、submit:値が送信される
 */
function setReadonlySelect(elm, on) {
    for (let i = 0; i < elm.options.length; i++) {
        let item = elm.options[i];
        if (on && !item.selected) {
            item.disabled = true;
        } else {
            item.disabled = false;
        }
    }
}