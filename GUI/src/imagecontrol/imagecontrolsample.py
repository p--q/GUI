#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.awt.ImageScaleMode import ISOTROPIC
from com.sun.star.awt import XActionListener
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK as ExecutableDialogResults_OK
from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
	def wrapper(*args, **kwargs):
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
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
		try:
			func(*args, **kwargs)  # Step Intoして中に入る。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return wrapper
# @enableRemoteDebugging
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。
	dialog, addControl = dialogCreator(ctx, smgr, {"Name": "ImageControlSample", "PositionX": 102, "PositionY": 41, "Width": 230, "Height": 151, "Title": "Image Control Sample", "Step": 0, "TabIndex": 0, "Moveable": True})  # "Sizeable": Trueが効かない。
	addControl("FixedText", {"Name": "Headerlabel", "PositionX": 6, "PositionY": 6, "Width": 210, "Height": 17, "Label": "This code-sample demonstrates how to create an ImageControlSample within a dialog.", "MultiLine": True})
	pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
	uno_path = pathsubstservice.getSubstituteVariableValue("$(prog)")  # fileurlでprogramフォルダへのパスが返ってくる。
	fileurl = "{}/intro.png".format(uno_path)  # 画像ファイルへのfileurl
	imageurl = os.path.normpath(unohelper.fileUrlToSystemPath(fileurl))  # fileurlをシステム固有のパスに変換して正規化する。
	addControl("ImageControl", {"PositionX": 6, "PositionY": 29, "Width": 218, "Height": 76, "Border": 0, "ScaleImage": True, "ScaleMode": ISOTROPIC, "ImageURL": fileurl})  # "ScaleImage": Trueで画像が歪む。
	addControl("Edit", {"Name": "EditFilePath","PositionX": 6, "PositionY": 111, "Width": 193, "Height": 14, "Text": imageurl})
	addControl("Button", {"Name": "ButtonFilePick", "PositionX": 199, "PositionY": 111, "Width": 25, "Height": 14, "Label": "~Browse", "PushButtonType": 0}, {"addActionListener": ActionListener(ctx, smgr, dialog)})  # PushButtonTypeの値はEnumではエラーになる。
	addControl("Button", {"PositionX": 90, "PositionY": 131, "Width": 50, "Height": 14, "Label": "~Close dialog", "PushButtonType": 1})  # PushButtonTypeの値はEnumではエラーになる。
	dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。
	# ノンモダルダイアログにするとき。
# 	showModelessly(ctx, smgr, docframe, dialog)
	# モダルダイアログにする。フレームに追加するとエラーになる。
	dialog.execute()
	dialog.dispose()
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, ctx, smgr, dialog):
		filters = {'WordPerfect Graphics': '*.wpg', 'SVM - StarView Meta File': '*.svm', 'PSD - Adobe Photoshop': '*.psd', 'EMF - Enhanced Meta File': '*.emf', 'PCD - Photo CD Base16': '*.pcd', 'PCD - Photo CD Base': '*.pcd', 'SGF - StarWriter SGF': '*.sgf', 'PGM - Portable Graymap': '*.pgm', 'SVG - Scalable Vector Graphics': '*.svg;*.svgz', 'PPM - Portable Pixelmap': '*.ppm', 'XBM - X Bitmap': '*.xbm', 'PBM - Portable Bitmap': '*.pbm', 'RAS - Sun Raster Image': '*.ras', 'WMF - Windows Metafile': '*.wmf', 'PCD - Photo CD Base4': '*.pcd', 'TGA - Truevision Targa': '*.tga', 'GIF - Graphics Interchange': '*.gif', 'Corel Presentation Exchange': '*.cmx', 'Adobe/Macromedia Freehand': '*.fh;*.fh1;*.fh2;*.fh3;*.fh4;*.fh5;*.fh6;*.fh7;*.fh8;*.fh9;*.fh10;*.fh11', 'CGM - Computer Graphics Metafile': '*.cgm', 'XPM - X PixMap': '*.xpm', 'MET - OS/2 Metafile': '*.met', 'DXF - AutoCAD Interchange Format': '*.dxf', 'JPEG - Joint Photographic Experts Group': '*.jpg;*.jpeg;*.jfif;*.jif;*.jpe', 'TIFF - Tagged Image File Format': '*.tif;*.tiff', 'PNG - Portable Network Graphic': '*.png', 'PCT - Mac Pict': '*.pct;*.pict', 'EPS - Encapsulated PostScript': '*.eps', 'BMP - Windows Bitmap': '*.bmp', 'PCX - Zsoft Paintbrush': '*.pcx'}  # 画像フィルターの辞書。
		filterall = "All Image Files"  # デフォルトで表示するフィルター名。
		template = FILEOPEN_SIMPLE
		try:  # 使えるのならFILEOPEN_PREVIEWを使う。
			from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_PREVIEW  # LibreOffice 5.3以上のみ
			template = FILEOPEN_PREVIEW
		except ImportError:
			pass
		filepicker = smgr.createInstanceWithArgumentsAndContext("com.sun.star.ui.dialogs.FilePicker", (template,), ctx)
		filepicker.appendFilter("All Files", "*.*")  # すべてのファイルを表示させるフィルターを最初に追加。
		filepicker.appendFilter(filterall, ";".join(filters.values()))  # すべての画像ファイルを表示させるフィルターを2番目に追加。
		[filepicker.appendFilter(key, filters[key]) for key in sorted(filters.keys())]  # フィルターは追加された順に表示されるのでfiltersをキーでソートしてから追加している。
		filepicker.setCurrentFilter(filterall)  # デフォルトで表示するフィルター名を設定。
		filepicker.setTitle("Insert Image")
		self.filepicker = filepicker
		self.workurl = ctx.getByName('/singletons/com.sun.star.util.thePathSettings').getPropertyValue("Work")  # Ubuntuではホームフォルダ、Windows10ではドキュメントフォルダのfileurlが返る。
		self.editcontrol = dialog.getControl("EditFilePath")
		self.imagecontrolmodel = dialog.getControl("ImageControl1").getModel()
		self.simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)
#	 @enableRemoteDebugging
	def actionPerformed(self, actionevent):
		dummy_control, dummy_controlmodel, name = eventSource(actionevent)
		if name == "ButtonFilePick":
			systempath = self.editcontrol.getText().strip()  # Editコントロールのテキストを取得。システム固有形式のパスが入っているはず。
			if os.path.exists(systempath):  # パスが実存するとき
				if os.path.isfile(systempath):  # ファイルへのパスであればその親フォルダのパスを取得する。
					systempath = os.path.dirname(systempath)
# 					self.filepicker.setDefaultName(os.path.basename(systempath))  # デフォルトファイル名を表示させたいが動かない。
				fileurl = unohelper.systemPathToFileUrl(systempath)  # fileurlに変換する。
			else:
				fileurl = self.workurl  # 実存するパスが取得できない時はホームフォルダのfileurlを取得。
			self.filepicker.setDisplayDirectory(fileurl)  # 表示するフォルダを設定。設定しないと「最近開いたファイル」が表示される。
			if self.filepicker.execute() == ExecutableDialogResults_OK:  # ファイル選択ダイアログを表示し、そのOKボタンがクリックされた時。
				fileurl = self.filepicker.getFiles()[0]  # ダイアログで選択されたファイルのパスを取得。fileurlのタプルで返ってくるので先頭の要素を取得。
				if self.simplefileaccess.exists(fileurl):  # fileurlが実存するとき
					self.imagecontrolmodel.setPropertyValue("ImageURL", fileurl)  # Imageコントロールに設定。
					systempath = unohelper.fileUrlToSystemPath(fileurl)  # fileurlをシステム固有形式に変換。
					self.editcontrol.setText(systempath)  # Editコントロールに表示。
	def disposing(self, eventobject):
		pass
def eventSource(event):  # イベントからコントロール、コントロールモデル、コントロール名を取得。
	control = event.Source  # イベントを駆動したコントロールを取得。
	controlmodel = control.getModel()  # コントロールモデルを取得。
	name = controlmodel.getPropertyValue("Name")  # コントロール名を取得。
	return control, controlmodel, name
def showModelessly(ctx, smgr, parentframe, dialog):  # ノンモダルダイアログにする。オートメーションではリスナー動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。
	frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しいフレームを生成。
	frame.initialize(dialog.getPeer())  # フレームにコンテナウィンドウを入れる。
	frame.setName(dialog.getModel().getPropertyValue("Name"))  # フレーム名をダイアログモデル名から取得（一致させる必要性はない）して設定。
	parentframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。
	dialog.setVisible(True)  # ダイアログを見えるようにする。
	return frame  # フレームにリスナーをつけるときのためにフレームを返す。
def dialogCreator(ctx, smgr, dialogprops):  # ダイアログと、それにコントロールを追加する関数を返す。まずダイアログモデルのプロパティを取得。
	dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)  # ダイアログの生成。
	if "PosSize" in dialogprops:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
		dialog.setPosSize(dialogprops.pop("PositionX"), dialogprops.pop("PositionY"), dialogprops.pop("Width"), dialogprops.pop("Height"), dialogprops.pop("PosSize"))  # ダイアログモデルのプロパティで設定すると単位がMapAppになってしまうのでコントロールに設定。
	dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログモデルの生成。
	dialogmodel.setPropertyValues(tuple(dialogprops.keys()), tuple(dialogprops.values()))  # ダイアログモデルのプロパティを設定。
	dialog.setModel(dialogmodel)  # ダイアログにダイアログモデルを設定。
	dialog.setVisible(False)  # 描画中のものを表示しない。
	def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
		control = None
		items, currentitemid = None, None
		if controltype == "Roadmap":  # Roadmapコントロールのとき、Itemsはダイアログモデルに追加してから設定する。そのときはCurrentItemIDもあとで設定する。
			if "Items" in props:  # Itemsはダイアログモデルに追加されてから設定する。
				items = props.pop("Items")
				if "CurrentItemID" in props:  # CurrentItemIDはItemsを追加されてから設定する。
					currentitemid = props.pop("CurrentItemID")
		if "PosSize" in props:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
			control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
			control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
			dialog.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
		else:  # Map AppFont (ma)のときはダイアログモデルにモデルを追加しないと正しくピクセルに変換されない。
			controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
			dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
		if items is not None:  # コントロールに追加されたRoadmapモデルにしかRoadmapアイテムは追加できない。
			for i, j in enumerate(items):  # 各Roadmapアイテムについて
				item = controlmodel.createInstance()
				item.setPropertyValues(("Label", "Enabled"), j)
				controlmodel.insertByIndex(i, item)  # IDは0から整数が自動追加される
			if currentitemid is not None:  #Roadmapアイテムを追加するとそれがCurrentItemIDになるので、Roadmapアイテムを追加してからCurrentIDを設定する。
				controlmodel.setPropertyValue("CurrentItemID", currentitemid)
		if control is None:  # コントロールがまだインスタンス化されていないとき
			control = dialog.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
		if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
			for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
				if val is None:
					getattr(control, key)()
				else:
					getattr(control, key)(val)
		return control  # 追加したコントロールを返す。
	def _createControlModel(controltype, props):  # コントロールモデルの生成。
		if not "Name" in props:
			props["Name"] = _generateSequentialName(controltype)  # Nameがpropsになければ通し番号名を生成。
		controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		if props:
			values = props.values()  # プロパティの値がタプルの時にsetProperties()でエラーが出るのでその対応が必要。
			if any(map(isinstance, values, [tuple]*len(values))):
				[setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
			else:
				controlmodel.setPropertyValues(tuple(props.keys()), tuple(values))
		return controlmodel
	def _generateSequentialName(controltype):  # コントロールの連番名の作成。
		i = 1
		flg = True
		while flg:
			name = "{}{}".format(controltype, i)
			flg = dialog.getControl(name)  # 同名のコントロールの有無を判断。
			i += 1
		return name
	return dialog, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
	from functools import wraps
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
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	if not hasattr(doc, "getCurrentController"):  # ドキュメント以外のとき。スタート画面も除外。
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/swriter", "_blank", 0, ())  # Writerのドキュメントを開く。
		while doc is None:  # ドキュメントのロード待ち。
			doc = XSCRIPTCONTEXT.getDocument()
	macro()
