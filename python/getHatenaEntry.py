import bs4
import re
import time
import requests

# 設定 > 詳細設定 > AtomPub に書かれている値を設定
hatena_id = "boppan"
blog_id = "boppan.hatenablog.com"
password = "********"

# コレクションURIを取得する関数
def get_collection_uri(hatena_id, blog_id, password):
  service_doc_uri = "https://blog.hatena.ne.jp/{hatena_id:}/{blog_id:}/atom".format(hatena_id=hatena_id, blog_id=blog_id)
  # リクエストを送る
  res_service_doc = requests.get(url=service_doc_uri, auth=(hatena_id, password))
  # 問題なければコレクションURIを返す
  if res_service_doc.ok:
    soup_servicedoc_xml = bs4.BeautifulSoup(res_service_doc.content, features="xml")
    collection_uri = soup_servicedoc_xml.collection.get("href")
    return collection_uri
  return False

# コレクションURIの取得
collection_uri = get_collection_uri(hatena_id, blog_id, password)

# ループ
MAX_ITERATER_NUM = 50
for i in range(MAX_ITERATER_NUM):

  # Basic認証で記事一覧を取得
  res_collection = requests.get(collection_uri, auth=(hatena_id, password))
  if not res_collection.ok:
    print("faild")
    continue

  # Beatifulsoup4でDOM化
  soup_collectino_xml = bs4.BeautifulSoup(res_collection.content, features="xml")
  # 記事のノードリスト取得
  entries = soup_collectino_xml.find_all("entry")

  # 記事IDとタイトルを取得して出力
  for e in entries:
    id = re.search(r"-(\d+)$", string=e.id.string).group(1)
    titlestr = e.title.string
    print(id + "：" + titlestr)

  # next
  link_next = soup_collectino_xml.find("link", rel="next")
  if not link_next:
    break
  collection_uri = link_next.get("href")
  if not collection_uri:
    break
  time.sleep(0.01)# 10ms