#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from xml.etree import ElementTree
from collections import ChainMap 
from com.sun.star.sheet import CellFlags as cf # 定数
def macro(documentevent=None):  # 引数はイベント駆動用。
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	filepath = "/opt/libreoffice5.2/share/registry/res/registry_ja.xcd"  # xmlファイルへのパス。
	xpath = './/node[@oor:name=".uno:FormatCellDialog"]'  # XPath。1つのノードだけ選択する条件にしないといけない。
	namespaces = {"oor": "{http://openoffice.org/2001/registry}",\
				"xs": "{http://www.w3.org/2001/XMLSchema}",\
				"xsi": "{http://www.w3.org/2001/XMLSchema-instance}"}  # 名前空間の辞書。replace()で置換するのに使う。
	traceToRoot(filepath, xpath, namespaces, doc)
def traceToRoot(filepath, xpath, namespaces, doc):  # xpathは子ノードを取得するXPath。1つのノードだけ選択する条件にしないといけない。
	tree = ElementTree.parse(filepath)  # xmlデータからElementTreeオブジェクト(xml.etree.ElementTree.ElementTree)を取得する。ElementTree.parse()のElementTreeはオブジェクト名ではなくてモジュール名。
	replaceWithValue, replaceWithKey = createReplaceFunc(namespaces)
	xpath = replaceWithValue(xpath)  # 名前空間の辞書のキーを値に変換。
	nodes = tree.findall(xpath)  # 起点となる子ノードを取得。
	outputs = []
	maps = []  # 子ノードをキー、親ノードの値とする辞書のリスト。
	parentnodes = True
	while parentnodes:  # 親ノードのリストの要素があるときTrue。
		xpath = "/".join([xpath, ".."])  # 親のノードのxpathの取得。
		parentnodes = tree.findall(xpath)  # 親ノードのリストを取得。
		maps.append({c:p for p in parentnodes for c in p})  # 新たな辞書をリストに追加。
	parentmap = ChainMap(*maps)  # 辞書をChainMapにする。
	for c in nodes:  # 各子ノードについて。
		outputs.append("\n{}".format(replaceWithKey(formatNode(c))))  # 名前空間の辞書の値をキーに変換して出力する。
		while c in parentmap:  # 親ノードが存在する時。
			c = parentmap[c]  # 親ノードを取得。
			outputs.append(replaceWithKey(formatNode(c)))  # 親ノードを出力。
	if outputs:
		datarows = [(i,) for i in outputs]  # Calcに出力するために行のリストにする。
		controller = doc.getCurrentController()  # コントローラーを取得。
		sheet = controller.getActiveSheet()  # アクティブなシートを取得。
		sheet.clearContents(cf.VALUE+cf.DATETIME+cf.STRING+cf.ANNOTATION+cf.FORMULA+cf.HARDATTR+cf.STYLES)  # セルの内容を削除。cf.HARDATTR+cf.STYLESでセル結合も解除。
		sheet[:len(datarows), :len(datarows[0])].setDataArray(datarows)  # シートに結果を出力する。
		cellcursor = sheet.createCursor()  # シート全体のセルカーサーを取得。
		cellcursor.gotoEndOfUsedArea(True)  # 使用範囲の右下のセルまでにセルカーサーのセル範囲を変更する。
		cellcursor.getColumns().setPropertyValue("OptimalWidth", True)  # セルカーサーのセル範囲の列幅を最適化する。
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
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
