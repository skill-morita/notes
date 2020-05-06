/**
 * �Q�l�y�[�W
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

/* charset �̒l�̗�:
 *	_autodetect, euc-jp, iso-2022-jp, shift_jis, unicode, utf-8,...
 */

/* filename: �ǂݍ��ރt�@�C���̃p�X
 * charset:	�����R�[�h
 * �߂�l:	 ������
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

/* filename: �ǂݍ��ރt�@�C���̃p�X
 * charset:	�����R�[�h
 * �߂�l:	 �s�P�ʂ̕�����̔z��
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
 * �e�L�X�g�o��
 * @param {*} filename  �����o���t�@�C���̃p�X
 * @param {*} text �o�͓��e
 * @param {*} charset �����R�[�h
 * @param {*} overwrite �㏑�ۑ�
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