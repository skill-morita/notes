// オプション
var OPTION_FLG = 2 + 16;
// 文字列リストファイルのパス
var FILEPATH = "C:/temp/wordlist.txt";
// 置換前文字列の列インデックス
var PRE_COL = 0;
// 置換後文字列の列インデックス
var POS_COL = 1;
// --------------------------------------------

// オブジェクト準備
var editor = Editor;

main();

// オブジェクト破棄
editor = null;

// --------------------------------------------
/**
 * 現在開いているタブの全文字列を、文字列リストで置換する
 * @returns 成否
 */
function main() {
	try {
		// 単語リスト読込
		var wordList = ReadTextFile(FILEPATH);
		if (wordList === null) {
			return false;
		}
		
		// 置換前文字列の長さで並び替え降順
		//wordList.sort(function(a, b) {
		//	if (a[PRE_COL].length > b[POS_COL].length) return -1;
		//	if (a[PRE_COL].length < b[POS_COL].length) return 1;
		//	return 0;
		//});

		// 順に一括置換
		for (var i = 0; i < wordList.length; i++) {
			editor.ReplaceAll(wordList[i][PRE_COL], wordList[i][POS_COL], OPTION_FLG);
		}

		// 再描画
		editor.ReDraw(0);
	} catch (error) {
		Messagebox("ERROR:" + error, 0);
		return false;
	}
}

/**
 * ファイル読取
 * @param {String} filepath ファイパス
 * @returns {String[][]} 読み取り内容
 */
function ReadTextFile(filepath) {
	var fileSysObj = new ActiveXObject("Scripting.FileSystemObject");

	// ファイルを開く
	var fileObj = fileSysObj.OpenTextFile(filepath, 1);
	if (fileObj.AtEndOfStream) {
		fileObj.Close();
		Messagebox("ERROR:置換リストにテキストがありません", 0);
		return null;
	}
Messagebox("a");

	// テキスト読取
	var text = fileObj.ReadAll();
	fileObj.Close();
	
	// データを配列に加工
	var res = [];
	var lines = text.split("\r\n");
	for (var i = 0; i < lines.length; i++) {
		if(lines[i].length > 0){
				res.push(lines[i].split(","));
		}
	}

	return res;
}