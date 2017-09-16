#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import sys
from com.sun.star.beans import PropertyValue
from com.sun.star.sheet.CellFlags import VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK_CANCEL
from com.sun.star.awt.MessageBoxType import QUERYBOX
from com.sun.star.awt.MessageBoxResults import CANCEL as MessageBoxResults_CANCEL
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
	if __name__ == "__main__":  # オートメーションのときはデバッグサーバーに接続しない。
		return func
	def wrapper(*args, **kwargs):
		import pydevd
		frame = None
		doc = XSCRIPTCONTEXT.getDocument()
		if doc:  # ドキュメントが取得できた時
			frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。
		else:
			currentframe = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()  # モードレスダイアログのときはドキュメントが取得できないので、モードレスダイアログのフレームからCreatorのフレームを取得する。
			frame = currentframe.getCreator()
		if frame:   
			import time
			indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
			maxrange = 2  # ステータスバーに表示するプログレスバーの目盛りの最大値。2秒ロスするが他に適当な告知手段が思いつかない。
			indicator.start("Trying to connect to the PyDev Debug Server for about 20 seconds.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
			t = 1  # プレグレスバーの初期値。
			while t<=maxrange:  # プログレスバーの最大値以下の間。
				indicator.setValue(t)  # プレグレスバーの位置を設定。
				time.sleep(1)  # 1秒待つ。
				t += 1  # プログレスバーの目盛りを増やす。
			indicator.end()  # reset()の前にend()しておかないと元に戻らない。
			indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。
		pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。
		try:
			func(*args, **kwargs)  # Step Intoして中に入る。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return wrapper
# @enableRemoteDebugging
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントモデルを取得。
	if not doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):  # Calcドキュメントに結果を出力するのでCalcドキュメントであることを確認する。
		raise RuntimeError("Please execute this macro with a Calc document.")
	sheets = doc.getSheets()  # シート列を取得。
	sheet = sheets[0]  # 最初のシートを取得。	
	if sheet.queryContentCells(VALUE+DATETIME+STRING):  # シートにデータがあるとき
		docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
		docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
		toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。 
		msgbox = toolkit.createMessageBox(docwindow, QUERYBOX, BUTTONS_OK_CANCEL, "Overwrite", 'The data already exists in "{}".\nDo you want to replace it?'.format(sheet.Name))
		if msgbox.execute()==MessageBoxResults_CANCEL:  # メッセージボックスを表示してキャンセルボタンがクリックされたとき。
			sys.exit()  # 終了する。
		msgbox.dispose()  # メッセージボックスを破棄。
	sheet.clearContents(VALUE+DATETIME+STRING+ANNOTATION+FORMULA+HARDATTR)  # シートの全内容をクリア。	
	configreader = createConfigReader(ctx, smgr)  # 読み込み専用の関数を取得。
	root = configreader("/org.openoffice.TypeDetection.Types/Types")  # org.openoffice.TypeDetectionパンケージのTypesコンポーネントのTypesノードを根ノードにする。
	props = "UIName", "Extensions", "MediaType"  # 取得するプロパティ名のタプル。
	filters = []  # フィルターのタプル。
	for childname in root.getElementNames():  # 子ノードの名前のタプルを取得。ノードオブジェクトの直接取得はできない模様。
		uiname, extensions, mediatype = root.getByName(childname).getPropertyValues(props)  # extentionsはタプルで返ってくる。
		extensions = "*.{}".format(";*.".join(extensions))
		filters.append((uiname, extensions, mediatype))
	r = len(filters)
	c = len(filters[0])
	cellrange = sheet[:r,:c]  # 代入するセルの範囲を指定する。
	cellrange.setDataArray(filters)
	cellrange.getColumns().OptimalWidth = True  # セルの幅を最適化する。実行時に入っているデータに合わせる。
	root.dispose() # ConfigurationAccessサービスのインスタンスを破棄。
def createConfigReader(ctx, smgr):  # ConfigurationProviderサービスのインスタンスを受け取る高階関数。
	configurationprovider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)  # ConfigurationProviderの取得。
	def configReader(path):  # ConfigurationAccessサービスのインスタンスを返す関数。
		node = PropertyValue(Name="nodepath", Value=path)
		return configurationprovider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	return configReader
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
	from functools import wraps
# 	import sys
#	 from com.sun.star.beans import PropertyValue
	from com.sun.star.script.provider import XScriptContext  
	def connectOffice(func):  # funcの前後でOffice接続の処理
		@wraps(func)
		def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
			try:
				ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
			except:
				print("Could not establish a connection with a running office.", file=sys.stderr)
				sys.exit()
			print("Connected to a running office ...")
			smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
			print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
			return func(ctx, smgr)  # 引数の関数の実行。
		def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
			cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
			node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
			ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
			return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
		return wrapper
	@connectOffice  # mainの引数にctxとsmgrを渡すデコレータ。
	def main(ctx, smgr):  # XSCRIPTCONTEXTを生成。
		class ScriptContext(unohelper.Base, XScriptContext):
			def __init__(self, ctx):
				self.ctx = ctx
			def getComponentContext(self):
				return self.ctx
			def getDesktop(self):
				return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
			def getDocument(self):
				return self.getDesktop().getCurrentComponent()
		return ScriptContext(ctx)  
	XSCRIPTCONTEXT = main()  # XSCRIPTCONTEXTを取得。
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
	doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
# 	doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
	if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
	flg = True
	while flg:
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		if doc is not None:
			flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
	macro()
	