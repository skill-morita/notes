/**
 * 参考ページ
 * http://aok.blue.coocan.jp/jscript/adodb.html
 */

/* StreamTypeEnum Values
 */
var adTypeBinary = 1;
var adTypeText = 2;

/* LineSeparatorEnum Values
 */
var adLF = 10;
var adCR = 13;
var adCRLF = -1;

/* StreamWriteEnum Values
 */
var adWriteChar = 0;
var adWriteLine = 1;

/* StreamReadEnum Values
 */
var adReadAll = -1;
var adReadLine = -2;

/* charset の値の例:
 *	_autodetect, euc-jp, iso-2022-jp, shift_jis, unicode, utf-8,...
 */

/* filename: 読み込むファイルのパス
 * charset:	文字コード
 * 戻り値:	 文字列
 */
function adoLoadText(filename, charset) {
  var stream;
  var text;
  stream = new ActiveXObject("ADODB.Stream");
  stream.type = adTypeText;
  stream.charset = charset;
  stream.open();
  stream.loadFromFile(filename);
  text = stream.readText(adReadAll);
  stream.close();
  return text;
}

/* filename: 読み込むファイルのパス
 * charset:	文字コード
 * 戻り値:	 行単位の文字列の配列
 */
function adoLoadLinesOfText(filename, charset) {
  var stream;
  var lines = new Array();
  stream = new ActiveXObject("ADODB.Stream");
  stream.type = adTypeText;
  stream.charset = charset;
  stream.open();
  stream.loadFromFile(filename);
  while (!stream.EOS) {
    lines.push(stream.readText(adReadLine));
  }
  stream.close();
  return lines;
}

/**
 * テキスト出力
 * @param {*} filename  書き出すファイルのパス
 * @param {*} text 出力内容
 * @param {*} charset 文字コード
 * @param {*} overwrite 上書保存
 */
function adoSaveText(filename, text, charset, overwrite) {
  var adSaveCreateNotExist = 1;
  var adSaveCreateOverWrite = 2;

  var overWriteFlg;
  if (overwrite) {
    overWriteFlg = adSaveCreateOverWrite;
  } else {
    overWriteFlg = adSaveCreateNotExist;
  }

  var stream = new ActiveXObject("ADODB.Stream");
  stream.type = adTypeText;
  stream.charset = charset;
  stream.open();
  stream.writeText(text);
  stream.saveToFile(filename, overWriteFlg);
  stream.close();
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
    alert("ERROR:置換リストにテキストがありません", 0);
    return null;
  }
  alert("a");

  // テキスト読取
  var text = fileObj.ReadAll();
  fileObj.Close();

  // データを配列に加工
  var res = [];
  var lines = text.split("\r\n");
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].length > 0) {
      res.push(lines[i].split(","));
    }
  }

  return res;
}