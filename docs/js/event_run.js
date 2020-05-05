/** ======================================================
 * イベントの発火
 * ====================================================== */

/**
 * イベントハンドラのボタン設定
 * @param {string} btnid ボタンのID
 * @param {string} bgcolorcd 背景色カラーコード
 * @param {string} bdcolorcd 罫線カラーコード
 */
function setEventButton(btnid, bgcolorcd, bdcolorcd) {
    let btn = document.querySelector(btnid);
    // プロパティで設定
    btn.onclick = function () { changeBorderColor(bdcolorcd) };
    // イベントリスナーで設定
    if (btn.addEventListener) {
        // IE8以下を除くブラウザ
        btn.addEventListener('click', { color: bgcolorcd, handleEvent: changeBgColor });
    } else if (btn.attachEvent) {
        // IE8以下
        btn.attachEvent('click', changeBgColor);
    }
}

/**
 * 背景色の変更
 * @param {any} e
 */
function changeBgColor(e) {
    document.querySelector('#colorChange').style.backgroundColor = this.color;
}

/**
 * 罫線の変更
 * @param {any} e
 */
function changeBorderColor(color) {
    document.querySelector('#colorChange').style.border = "solid 3px " + color;
}

// -------------------------------------------------------
// イベントの設定
setEventButton('#btn1', '#4682b4', '#000080');
setEventButton('#btn2', '#ff6347', '#dc143c');
setEventButton('#btn3', '#008b8b', '#2f4f4f');

// 間接的にイベント設定
document.querySelector('#btn1_1').onclick = function () {
    // (1)イベントリスナーを使用する
    let btn = document.querySelector('#btn1');
    if (btn.addEventListener) {
        // IE8以下を除くブラウザ
        // 方法1
        let event = new MouseEvent("click");
        btn.dispatchEvent(event);
        // // 方法2
        // let event2 = new Event("click");
        // btn.dispatchEvent(event2);
        // // 方法3
        // let event3 = document.createEvent("MouseEvents");
        // event3.initEvent("click", false, true);
        // btn.dispatchEvent(event3);
    } else if (btn.attachEvent) {
        // IE8以下
        btn.fireEvent("onclick");
    }
};

document.querySelector('#btn2_1').onclick = function () {
    // (2)イベントハンドラプロパティを使用する
    // onclickプロパティに設定したものは実行されるが、addEventListenerしたものは反応しない。
    let btn = document.querySelector('#btn2');
    btn.onclick();
};

document.querySelector('#btn3_1').onclick = function () {
    // (3)HTMLElementのclickメソッドを使用する
    let btn = document.querySelector('#btn3');
    btn.click();
};

