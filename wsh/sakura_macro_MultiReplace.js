/**
 * @file sakura_macro_MultiReplace.js
 * @author boppan
 * @description 
 * 複数単語の一括置換
 * 参考: [【サクラエディタ】文字列を一括で複数置換する \- Qiita](https://qiita.com/MasayaOkuno/items/ba6bec13f01cdd78d909)
 * 
 * 1. 置換単語リストをクリップボードにコピーする
 *     タブ区切りで左に置換前、右に置換後
 *    【形式】
 *      testA	test1
 *      testB	test2
 *      testC	test3
 *  2. 置換したいタブをアクティブにする
 *  3. マクロ実行
 */
// --------------------------------------------
// setting
// --------------------------------------------
// [サクラエディタの置換オプション \| You Look Too Cool](https://stabucky.com/wp/archives/4678)
// オプション 2:英大文字と小文字を区別する 16:置換ダイアログを自動的に閉じる
var OPTION_FLG = 2 + 16;
// --------------------------------------------
// function
// --------------------------------------------
/**
 * 現在選択中のタブの全文字列を、文字列リストで置換する
 */
function multiReplace(editor) {
	try {
		// 単語リスト取得
		var clip = editor.GetClipboard(0);
		var wordList = clip.split('\r\n');

		// 文字数で降順(testABとtestAがあった場合、testABが優先される)
		wordList.sort(function (a, b) {
			if (a.length > b.length) return -1;
			if (a.length < b.length) return 1;
			return 0;
		});

		// 順に一括置換
		for (var i = 0; i < wordList.length; i++) {
			var line = wordList[i].split('\t');
			var oldStr = line[0];
			var newStr = line[1];
			if (typeof oldStr === 'string' && typeof newStr === 'string') {
				// TODO 単語が無い時の高速化
				editor.ReplaceAll(oldStr, newStr, OPTION_FLG);
			}
		}

		// 再描画
		editor.ReDraw(0);
	} catch (error) {
		alert("ERROR:" + error);
	}
}
// --------------------------------------------
// main
// --------------------------------------------
// オブジェクト準備
var editor = Editor;

multiReplace(editor);

// オブジェクト破棄
editor = null;
