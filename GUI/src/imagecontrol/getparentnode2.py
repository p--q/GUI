# -*- coding: utf-8 -*-
from xml.etree import ElementTree
from collections import ChainMap 
def traceToRoot():
	tree = ElementTree.parse('/opt/libreoffice5.2/share/registry/main.xcd')  # xmlデータからElementTreeオブジェクト(xml.etree.ElementTree.ElementTree)を取得する。ElementTree.parse()のElementTreeはオブジェクト名ではなくてモジュール名。
# 	xpath = './/prop[@oor:name="StartCenterBackgroundColor"]'  # 子ノードを取得するXPath。1つのノードだけ選択する条件にしないといけない。
	xpath = './/prop[@oor:name="CaptionText"]'  # 子ノードを取得するXPath。複数の子ノードを返ってくる例。うまく動かない。
	namespaces = {"oor": "{http://openoffice.org/2001/registry}",\
			 "xs": "{http://www.w3.org/2001/XMLSchema}",\
			 "xsi": "{http://www.w3.org/2001/XMLSchema-instance}",\
			 "xml": "{http://www.w3.org/XML/1998/namespace}"}  # 名前空間の辞書。replace()で置換するのに使う。
	replaceWithValue, replaceWithKey = createReplaceFunc(namespaces)
	xpath = replaceWithValue(xpath)  # 名前空間の辞書のキーを値に変換。
	nodes = tree.findall(xpath)  # 起点となる子ノードを取得。
	maps = []  # 子ノードをキー、親ノードの値とする辞書のリスト。
	parentnodes = True
	while parentnodes:  # 親ノードのリストの要素があるときTrue。
		xpath = "/".join([xpath, ".."])  # 親のノードのxpathの取得。
		parentnodes = tree.findall(xpath)  # 親ノードのリストを取得。
		maps.append({c:p for p in parentnodes for c in p})  # 新たな辞書をリストに追加。
	parentmap = ChainMap(*maps)  # 辞書をChainMapにする。
	for c in nodes:  # 各子ノードについて。
		print("\n{}".format(replaceWithKey(formatNode(c))))  # 名前空間の辞書の値をキーに変換して出力する。
		while c in parentmap:  # 親ノードが存在する時。
			c = parentmap[c]  # 親ノードを取得。
			print(replaceWithKey(formatNode(c)))  # 親ノードを出力。
def formatNode(node):  # 引数はElement オブジェクト。タグ名と属性を出力する。属性の順番は保障されない。
	tag = node.tag  # タグ名を取得。
	attribs = []  # 属性をいれるリスト。
	for key, val in node.items():  # ノードの各属性について。
		attribs.append('{}="{}"'.format(key, val))  # =で結合。
	attrib = " ".join(attribs)  # すべての属性を結合。
	n = "{} {}".format(tag, attrib) if attrib else tag  # タグ名と属性を結合する。
	return "<{}>".format(n)					
def createReplaceFunc(namespaces):  # 引数はキー名前空間名、値は名前空間を波括弧がくくった文字列、の辞書。
	def replaceWithValue(txt):  # 名前空間の辞書のキーを値に置換する。
		for key, val in namespaces.items():
			txt = txt.replace("{}:".format(key), val)
		return txt
	def replaceWithKey(txt):  # 名前空間の辞書の値をキーに置換する。
		for key, val in namespaces.items():
			txt = txt.replace(val, "{}:".format(key))
		return txt
	return replaceWithValue, replaceWithKey
if __name__ == "__main__": 
	traceToRoot()
	