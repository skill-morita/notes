# ----------------------------------
# Wikiの国名漢字一覧を取得してみる
# ----------------------------------
import requests
from bs4 import BeautifulSoup
import lxml.html

def output(filename, contents):
    path_w = '_temp/' + filename
    with open(path_w, mode='w', encoding='utf8') as f:
        for line in contents:
            f.write(line + '\n')

    with open(path_w, encoding='utf8') as f:
        print(f.read())

def get_contents(_params):
    res = requests.get(url=URL, params=_params)

    # アクセスしたURLとレスポンスの型取得
    print(res.url)
    # print(res.headers['Content-Type'])

    # jsonに変換
    json_obj = res.json()
    # print(json_obj)

    return json_obj

URL = 'http://ja.wikipedia.org/w/api.php'

# 国
PARAMS = {
    "action": "parse",
    "page": "国名の漢字表記一覧",
    "format": "json"
}

json_obj = get_contents(PARAMS)
# ページ内容取得
title = json_obj['parse']['title']
html_str = json_obj['parse']['text']['*']
# print(html_str)

# ノード取得
dom = lxml.html.fromstring(html_str)
elmList = dom.xpath(u"//tr/td[1]/a[2]")
contents = []
for elm in elmList:
    # 次のtd内の文字列
    td = elm.getparent().getnext()
    nation = elm.text
    kanji = td.text
    if kanji is None:
        continue
    contents.append(nation + "\t" + kanji)

output("getWiki_" + title + ".txt", contents)
