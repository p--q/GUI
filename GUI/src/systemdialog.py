#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import sys
from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE, FILEOPEN_LINK_PREVIEW_IMAGE_TEMPLATE, FILEOPEN_PLAY,FILEOPEN_READONLY_VERSION, FILEOPEN_LINK_PREVIEW, FILESAVE_SIMPLE, FILESAVE_AUTOEXTENSION_PASSWORD, FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS, FILESAVE_AUTOEXTENSION_SELECTION, FILESAVE_AUTOEXTENSION_TEMPLATE, FILESAVE_AUTOEXTENSION
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
def macro():  # オートメーションでFilePickerサービスをインスタンス化するとクラッシュする。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	templateurl = ctx.getByName('/singletons/com.sun.star.util.thePathSettings').getPropertyValue("Work")  # デフォルトで表示するフォルダを取得。
	filters = {'WordPerfect Graphics': '*.wpg', 'SVM - StarView Meta File': '*.svm', 'PSD - Adobe Photoshop': '*.psd', 'EMF - Enhanced Meta File': '*.emf', 'PCD - Photo CD Base16': '*.pcd', 'PCD - Photo CD Base': '*.pcd', 'SGF - StarWriter SGF': '*.sgf', 'PGM - Portable Graymap': '*.pgm', 'SVG - Scalable Vector Graphics': '*.svg;*.svgz', 'PPM - Portable Pixelmap': '*.ppm', 'XBM - X Bitmap': '*.xbm', 'PBM - Portable Bitmap': '*.pbm', 'RAS - Sun Raster Image': '*.ras', 'WMF - Windows Metafile': '*.wmf', 'PCD - Photo CD Base4': '*.pcd', 'TGA - Truevision Targa': '*.tga', 'GIF - Graphics Interchange': '*.gif', 'Corel Presentation Exchange': '*.cmx', 'Adobe/Macromedia Freehand': '*.fh;*.fh1;*.fh2;*.fh3;*.fh4;*.fh5;*.fh6;*.fh7;*.fh8;*.fh9;*.fh10;*.fh11', 'CGM - Computer Graphics Metafile': '*.cgm', 'XPM - X PixMap': '*.xpm', 'MET - OS/2 Metafile': '*.met', 'DXF - AutoCAD Interchange Format': '*.dxf', 'JPEG - Joint Photographic Experts Group': '*.jpg;*.jpeg;*.jfif;*.jif;*.jpe', 'TIFF - Tagged Image File Format': '*.tif;*.tiff', 'PNG - Portable Network Graphic': '*.png', 'PCT - Mac Pict': '*.pct;*.pict', 'EPS - Encapsulated PostScript': '*.eps', 'BMP - Windows Bitmap': '*.bmp', 'PCX - Zsoft Paintbrush': '*.pcx'}  # 画像フィルターの辞書。
	filterall = "All Image Files"
	filters[filterall] = ";".join(filters.values())  # すべての画像ファイルをまとめたフィルターを辞書に追加。
	filters["All Files"] = "*.*"  # すべてのファイルのフィルターを辞書に追加。
	# ファイルを開くダイアログ。
	fileopens = FILEOPEN_SIMPLE, FILEOPEN_LINK_PREVIEW_IMAGE_TEMPLATE, FILEOPEN_PLAY, FILEOPEN_READONLY_VERSION, FILEOPEN_LINK_PREVIEW
	templatenames = "FILEOPEN_SIMPLE", "FILEOPEN_LINK_PREVIEW_IMAGE_TEMPLATE", "FILEOPEN_PLAY", "FILEOPEN_READONLY_VERSION", "FILEOPEN_LINK_PREVIEW"
	for template, templatename in zip(fileopens, templatenames):
		filepicker = createFilepicker(ctx, smgr, template)
		settingFilePicker(filepicker, filters, filterall, templateurl, templatename)
	# LibreOffice5.3以上でのみ使えるファイルを開くダイアログ。
	try:
		from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_PREVIEW, FILEOPEN_LINK_PLAY  # LibreOffice 5.3以上のみ
		fileopens = FILEOPEN_PREVIEW, FILEOPEN_LINK_PLAY
		templatenames = "FILEOPEN_PREVIEW", "FILEOPEN_LINK_PLAY"
		for template, templatename in zip(fileopens, templatenames):
			templatename = "{} since LibreOffice 5.3".format(templatename)
			filepicker = createFilepicker(ctx, smgr, template)
			settingFilePicker(filepicker, filters, filterall, templateurl, templatename)
	except ImportError:
		from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
		from com.sun.star.awt.MessageBoxType import WARNINGBOX
		doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
		docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
		docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
		toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。 
		msgbox = toolkit.createMessageBox(docwindow, WARNINGBOX, BUTTONS_OK, "Overwrite", "FILEOPEN_PREVIEW and FILEOPEN_LINK_PLAY need \nLibreOffice version more than 5.3.")
		msgbox.execute()
		msgbox.dispose()  # メッセージボックスを破棄。
	# ファイルを保存するダイアログ。
	filesaves = FILESAVE_SIMPLE, FILESAVE_AUTOEXTENSION_PASSWORD, FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS, FILESAVE_AUTOEXTENSION_SELECTION, FILESAVE_AUTOEXTENSION_TEMPLATE, FILESAVE_AUTOEXTENSION
	templatenames = "FILESAVE_SIMPLE", "FILESAVE_AUTOEXTENSION_PASSWORD", "FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS", "FILESAVE_AUTOEXTENSION_SELECTION", "FILESAVE_AUTOEXTENSION_TEMPLATE", "FILESAVE_AUTOEXTENSION"
	for template, templatename in zip(filesaves, templatenames):
		filepicker = createFilepicker(ctx, smgr, template)
		settingFilePicker(filepicker, filters, filterall, templateurl, templatename)		
	# フォルダの選択するダイアログ。	
	folderpicker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", ctx)
	folderpicker.setDisplayDirectory(templateurl)
	folderpicker.setTitle("FolderPicker")
	folderpicker.execute()
def createFilepicker(ctx, smgr, template):
	# サービスマネージャーのcreateInstanceWithArgumentsAndContext()でTemplateDescriptionを指定。
	filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (template,), ctx)
	# サービスマネージャーのinitialize()でTemplateDescriptionを指定。
# 	filepicker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", ctx)
# 	filepicker.initialize((template,))	
	return filepicker
def settingFilePicker(filepicker, filters, filterall, templateurl, templatename):
	[filepicker.appendFilter(key, filters[key]) for key in sorted(filters.keys())]  # フィルターは追加された順に表示されるのでfiltersをキーでソートしてから追加している。
	filepicker.setCurrentFilter(filterall)  # デフォルトで表示するフィルターを設定。linuxBeanのFILESAVE系では「すべての形式」(*以外のフィルターの拡張子を足したもの）というのが表示されてしまう。	
	filepicker.setDisplayDirectory(templateurl)  # デフォルトで表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
	filepicker.setDefaultName("UML.png")  # FILEOPEN系では動かない。Windows10では拡張子がfh11になる。		
	filepicker.setTitle(templatename)
	filepicker.execute()	
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
	from functools import wraps
# 	import sys
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
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	if not hasattr(doc, "getCurrentController"):  # ドキュメント以外のとき。スタート画面も除外。
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/swriter", "_blank", 0, ())  # Writerのドキュメントを開く。
		while doc is None:  # ドキュメントのロード待ち。
			doc = XSCRIPTCONTEXT.getDocument()
	macro()