#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.beans import PropertyValue
from com.sun.star.ui.dialogs.TemplateDescription import FILESAVE_AUTOEXTENSION
from com.sun.star.ui.dialogs.ExtendedFilePickerElementIds import CHECKBOX_AUTOEXTENSION
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK as ExecutableDialogResults_OK


def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
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
	thepathsettings = ctx.getByName('/singletons/com.sun.star.util.thePathSettings')
	filepicker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", ctx)
	filepicker.setDefaultName("MyExampleDocument")
	templateurl = thepathsettings.getPropertyValue("Work")  
	filepicker.setDisplayDirectory(templateurl)
	configurationprovider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)  # ConfigurationProviderの取得。
	configreader = createConfigReader(configurationprovider)  # 読み込み専用の関数を取得。
	root = configreader("/org.openoffice.Office.UI/FilterClassification/GlobalFilters/Classes")  # グローバルフィルター。
	props = "DisplayName", "Filters"  # 取得するプロパティ名のタプル。
	filters = {}  # フィルターの辞書。UINameをキー、拡張子のタプルを値とする。
	for childname in root.getElementNames():  # 子ノードの名前のタプルを取得。ノードオブジェクトの直接取得はできない模様。
		node = root.getByName(childname)  # ノードオブジェクトを取得。
		displayname, globalfilters = node.getPropertyValues(props)  # propの値を取得。displaynameは翻訳語のものが返ってくる。
		filters[displayname] = ";".join(globalfilters)
# 	filters[displayname] = globalfilters
# 	filterall = "All Files"
# 	filters[filterall] = "*.*"
# 	for key, val in filters.items():
# 		filepicker.appendFilter(key, val)
		
		
		
# 		filepicker.appendFilter(key, val)

	filepicker.appendFilter("All Files", "*.*")
	
# 	filepicker.setCurrentFilter(filterall)
	filepicker.appendFilter("OpenDocument Text Template", "writer8_template")
	filepicker.appendFilter("OpenDocument Text", "writer 8")
	
	
# 	filepicker.initialize((FILESAVE_AUTOEXTENSION,))

# 	filepicker.initialize((2,))

# 	filepicker.setValue(CHECKBOX_AUTOEXTENSION, 6, True)
	result = filepicker.execute()
# 	if result==ExecutableDialogResults_OK:
# 		pathlist = filepicker.getFiles()
# 		if pathlist:
# 			storepath = pathlist[0]
# 	workurl = thepathsettings.getPropertyValue("Work") 	
# 	folderpicker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", ctx)
# 	folderpicker.setDisplayDirectory(workurl)
# 	folderpicker.setTitle("My Title")
# 	result = folderpicker.execute()
# 	if result==ExecutableDialogResults_OK:
# 		returnfolder = folderpicker.getDirectory()
	
	

def createConfigReader(cp):  # ConfigurationProviderサービスのインスタンスを受け取る高階関数。
	def getRoot(path):  # ConfigurationAccessサービスのインスタンスを返す関数。
		node = PropertyValue(Name="nodepath", Value=path)
		return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
	return getRoot
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
	from functools import wraps
	import sys
# 	from com.sun.star.beans import PropertyValue
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
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	if not hasattr(doc, "getCurrentController"):  # ドキュメント以外のとき。スタート画面も除外。
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/swriter", "_blank", 0, ())  # Writerのドキュメントを開く。
		while doc is None:  # ドキュメントのロード待ち。
			doc = XSCRIPTCONTEXT.getDocument()
	macro()