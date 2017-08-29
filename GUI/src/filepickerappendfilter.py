#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import time
from functools import wraps
from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE
from com.sun.star.beans import StringPair
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
# 			import time
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
def timethis(func):
	doc = XSCRIPTCONTEXT.getDocument()  # Writerドキュメントを取得。
	@wraps(func)
	def wrapper(*args, **kwargs):
		start = time.perf_counter()
		r = func(*args, **kwargs)
		end = time.perf_counter()
		doctext = doc.getText().getString()  # すでにドキュメントにある文字列を取得。
		doc.getText().setString('{}\n{}.{} : {}'.format(doctext, func.__module__, func.__name__, end-start))  # Writerドキュメントに出力。 上書き。 
# 		print('{}.{} : {}'.format(func.__module__, func.__name__, end-start))
		return r
	return wrapper	
# @enableRemoteDebugging
def macro():  # オートメーションでFilePickerサービスをインスタンス化するとクラッシュする。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	templateurl = ctx.getByName('/singletons/com.sun.star.util.thePathSettings').getPropertyValue("Work")  # デフォルトで表示するフォルダを取得。
	filters = {'WordPerfect Graphics': '*.wpg', 'SVM - StarView Meta File': '*.svm', 'PSD - Adobe Photoshop': '*.psd', 'EMF - Enhanced Meta File': '*.emf', 'PCD - Photo CD Base16': '*.pcd', 'PCD - Photo CD Base': '*.pcd', 'SGF - StarWriter SGF': '*.sgf', 'PGM - Portable Graymap': '*.pgm', 'SVG - Scalable Vector Graphics': '*.svg;*.svgz', 'PPM - Portable Pixelmap': '*.ppm', 'XBM - X Bitmap': '*.xbm', 'PBM - Portable Bitmap': '*.pbm', 'RAS - Sun Raster Image': '*.ras', 'WMF - Windows Metafile': '*.wmf', 'PCD - Photo CD Base4': '*.pcd', 'TGA - Truevision Targa': '*.tga', 'GIF - Graphics Interchange': '*.gif', 'Corel Presentation Exchange': '*.cmx', 'Adobe/Macromedia Freehand': '*.fh;*.fh1;*.fh2;*.fh3;*.fh4;*.fh5;*.fh6;*.fh7;*.fh8;*.fh9;*.fh10;*.fh11', 'CGM - Computer Graphics Metafile': '*.cgm', 'XPM - X PixMap': '*.xpm', 'MET - OS/2 Metafile': '*.met', 'DXF - AutoCAD Interchange Format': '*.dxf', 'JPEG - Joint Photographic Experts Group': '*.jpg;*.jpeg;*.jfif;*.jif;*.jpe', 'TIFF - Tagged Image File Format': '*.tif;*.tiff', 'PNG - Portable Network Graphic': '*.png', 'PCT - Mac Pict': '*.pct;*.pict', 'EPS - Encapsulated PostScript': '*.eps', 'BMP - Windows Bitmap': '*.bmp', 'PCX - Zsoft Paintbrush': '*.pcx'}  # 画像フィルターの辞書。
	filterall = "All Image Files"
	filters[filterall] = ";".join(filters.values())  # すべての画像ファイルをまとめたフィルターを辞書に追加。
	filters["All Files"] = "*.*"  # すべてのファイルのフィルターを辞書に追加。
	filterpicker = appendFilterGroup(ctx, smgr, filters, filterall, templateurl)
	filterpicker.execute()
	filterpicker = appendFilter(ctx, smgr, filters, filterall, templateurl)
	filterpicker.execute()
@timethis
def appendFilterGroup(ctx, smgr, filters, filterall, templateurl):	
	filterpairs = []
	[filterpairs.append(StringPair(First=key, Second=filters[key])) for key in sorted(filters.keys())]
	filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (FILEOPEN_SIMPLE,), ctx)
	filepicker.setDisplayDirectory(templateurl)  # デフォルトで表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
	filepicker.appendFilterGroup("Filters", filterpairs)
	filepicker.setCurrentFilter(filterall)  # デフォルトで表示するフィルターを設定。linuxBeanのFILESAVE系では「すべての形式」(*以外のフィルターの拡張子を足したもの）というのが表示されてしまう。
	return filepicker
@timethis	
def appendFilter(ctx, smgr, filters, filterall, templateurl):	
	filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (FILEOPEN_SIMPLE,), ctx)
	filepicker.setDisplayDirectory(templateurl)  # デフォルトで表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
	[filepicker.appendFilter(key, filters[key]) for key in sorted(filters.keys())]  # フィルターは追加された順に表示されるのでfiltersをキーでソートしてから追加している。filepicker.setCurrentFilter(filterall)  # デフォルトで表示するフィルターを設定。linuxBeanのFILESAVE系では「すべての形式」(*以外のフィルターの拡張子を足したもの）というのが表示されてしまう。
	filepicker.setCurrentFilter(filterall)  # デフォルトで表示するフィルターを設定。linuxBeanのFILESAVE系では「すべての形式」(*以外のフィルターの拡張子を足したもの）というのが表示されてしまう。
	return filepicker
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
# 	from functools import wraps
	import sys
	from com.sun.star.beans import PropertyValue
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
# 	doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
	doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
	if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
	flg = True
	while flg:
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		if doc is not None:
			flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
	macro()