#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import Rectangle
from com.sun.star.awt import XMenuListener
from com.sun.star.awt import MenuItemStyle as mi  # 定数
from com.sun.star.awt import PopupMenuDirection  # 定数
from com.sun.star.util import XCloseListener
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。
	dialog, addControl = dialogCreator(ctx, smgr, {"PositionX": 102, "PositionY": 41, "Width": 200, "Height": 140, "Title": "Menu-Dialog", "Name": "Dialog1", "Step": 1, "TabIndex": 0, "Moveable": True})
	createMenu = menuCreator(ctx, smgr)
	menulistener = MenuListener(dialog)  # ポップアップメニューにつけるメニューリスナーを取得。
	items = ("First Entry", mi.CHECKABLE+mi.AUTOCHECK, {"checkItem": True}),\
			("First Radio Entry", mi.RADIOCHECK+mi.AUTOCHECK, {"enableItem": False}),\
			("Second Radio Entry", mi.RADIOCHECK+mi.AUTOCHECK),\
			("Third Radio Entry", mi.RADIOCHECK+mi.AUTOCHECK, {"checkItem": True}),\
			(),\
			("Fourth Entry", mi.CHECKABLE+mi.AUTOCHECK, {"checkItem": True}),\
			("Fifth Entry", mi.CHECKABLE+mi.AUTOCHECK),\
			("Sixth Entry", 0),\
			("~Close", 0, {"setCommand": "close"})
	popupmenu =  createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 右クリックでまず呼び出すポップアップメニュー。
	items = ("First Entry", mi.CHECKABLE+mi.AUTOCHECK, {"checkItem": True}),\
			("Second Entry", 0)
	subpopupmenu =  createMenu("PopupMenu", items, {"addMenuListener": menulistener})  # 入れ子にするポップアップメニュー。
	popupmenu.setPopupMenu (8, subpopupmenu)  # ポップアップメニューを入れ子にする。
	addControl("FixedText", {"Name": "Headerlabel", "PositionX": 6, "PositionY": 6, "Width": 200, "Height": 8, "Label": "This code-sample demonstrates the creation of a popup-menu."})
	mouselistener = MouseListener(ctx, smgr, popupmenu)
	fixedtext2 = addControl("FixedText", {"PositionX": 50, "PositionY": 50, "Width": 100, "Height": 8, "Label": "Right-click here"}, {"addMouseListener": mouselistener})
	dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。
	# ノンモダルダイアログにするとき。
# 	frame = showModelessly(ctx, smgr, docframe, dialog)
# 	menulistener.frame = frame  # メニューで閉じるためにフレームをメニューリスナーに渡す。
# 	args = popupmenu, subpopupmenu, menulistener, fixedtext2, mouselistener
# 	frame.addCloseListener(CloseListener(args))  # CloseListener
	# モダルダイアログにする。フレームに追加するとエラーになる。
	dialog.execute()
	popupmenu.removeMenuListener(menulistener)
	subpopupmenu.removeMenuListener(menulistener)
	dialog.dispose()
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログでのみ使用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):
		popupmenu, subpopupmenu, menulistener, fixedtext2, mouselistener = self.args
		popupmenu.removeMenuListener(menulistener)
		subpopupmenu.removeMenuListener(menulistener)
		fixedtext2.removeMouseListener(mouselistener)
		eventobject.Source.removeCloseListener(self)
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "getsownership: {}\nSource: {}".format(getsownership, eventobject.Source))	
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  # 発火しない。
		eventobject.Source.removeCloseListener(self)
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
class MouseListener(unohelper.Base, XMouseListener):  # Editコントロールではうまく動かない。
	def __init__(self, ctx, smgr, popupmenu):
		self.popupmenu = popupmenu
	def mousePressed(self, mouseevent):  # マウスがクリックされた時。
		control = mouseevent.Source  # イベントを駆動したコントロールを取得。
		name = control.getModel().getPropertyValue("Name")  # コントロール名を取得。
		if name == "FixedText1":  # コントロール名で限定。
			if mouseevent.PopupTrigger:  # 右クリックのとき
				pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
				self.popupmenu.execute(control.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向。ここでメソッドを出るらしい。
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):  # モダルダイアログを閉じる時に発火する。
		eventobject.Source.removeMouseListener(self)
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self, dialog):
		self.dialog = dialog
		self.frame = None  # ノンモダルダイアログで使用。
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。
		source = menuevent.Source
		cmd = source.getCommand(menuevent.MenuId)
		if cmd == "close":
			if self.frame is None:  # フレームがないときはモダルダイアログ。
				self.dialog.endExecute()  # ウィンドウを閉じる。モダルダイアログではこれだけで閉じる。 モードレスダイアログは閉じない。
			else:  # フレームがあるときはノンモダルダイアログ。
				self.frame.close(True)  # フレームを閉じる。self.dialog.dispose()でも閉じる。モダルダイアログではdispose()ではドキュメントまでも閉じてしまう。
	def itemActivated(self, menuevent):
		pass
	def itemDeactivated(self, menuevent):
		pass
	def disposing(self, eventobject):  # モダルダイアログを閉じる時でも発火しない。
		eventobject.Source.removeMenuListener(self)
		createLog(eventobject.Source, "{}_{}".format(self.__class__.__name__, inspect.currentframe().f_code.co_name), "Source: {}".format(eventobject.Source))	
def menuCreator(ctx, smgr):  #  メニューバーまたはポップアップメニューを作成する関数を返す。
	def createMenu(menutype, items, attr=None):  # menutypeはMenuBarまたはPopupMenu、itemsは各メニュー項目の項目名、スタイル、適用するメソッドのタプルのタプル、attrは各項目に適用する以外のメソッド。
		if attr is None:
			attr = {}
		menu = smgr.createInstanceWithContext("com.sun.star.awt.{}".format(menutype), ctx)
		for i, item in enumerate(items, start=1):  # 各メニュー項目について。
			if item:
				if len(item) > 2:  # タプルの要素が3以上のときは3番目の要素は適用するメソッドの辞書と考える。
					item = list(item)
					attr[i] = item.pop()  # メニュー項目のIDをキーとしてメソッド辞書に付け替える。
				menu.insertItem(i, *item, i-1)  # ItemId, Text, ItemSytle, ItemPos。ItemIdは1から始まり区切り線は含まない。ItemPosは0から始まり区切り線を含む。
			else:  # 空のタプルの時は区切り線と考える。
				menu.insertSeparator(i-1)  # ItemPos
		if attr:  # メソッドの適用。
			for key, val in attr.items():  # keyはメソッド名あるいはメニュー項目のID。
				if isinstance(val, dict):  # valが辞書の時はkeyは項目ID。valはcreateMenu()の引数のitemsであり、itemsの３番目の要素にキーをメソッド名とする辞書が入っている。
					for method, arg in val.items():  # 辞書valのキーはメソッド名、値はメソッドの引数。
						if method in ("checkItem", "enableItem", "setCommand", "setHelpCommand", "setHelpText", "setTipHelpText"):  # 第1引数にIDを必要するメソッド。
							getattr(menu, method)(key, arg)
						else:
							getattr(menu, method)(arg)
				else:
					getattr(menu, key)(val)
		return menu
	return createMenu
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
			if controltype=="Grid":
				control = smgr.createInstanceWithContext("com.sun.star.awt.grid.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
			else: 
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
		if controltype=="Grid":
			controlmodel = dialogmodel.createInstance("com.sun.star.awt.grid.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
		else: 
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
import os, inspect
from datetime import datetime
C = 100  # カウンターの初期値。
TIMESTAMP = datetime.now().isoformat().split(".")[0].replace("-", "").replace(":", "")  # コピー先ファイル名に使う年月日T時分秒を結合した文字列を取得。
def createLog(source, filename, txt):  # 年月日T時分秒リスナーのインスタンス名_メソッド名(_オプション).logファイルを作成。txtはファイルに書き込むテキスト。dirpathはファイルを書き出すディレクトリ。
	path = XSCRIPTCONTEXT.getDocument().getURL() if __file__.startswith("vnd.sun.star.tdoc:") else __file__  # このスクリプトのパス。fileurlで返ってくる。埋め込みマクロの時は埋め込んだドキュメントのURLで代用する。
	thisscriptpath = unohelper.fileUrlToSystemPath(path)  # fileurlをsystempathに変換。
	dirpath = os.path.dirname(thisscriptpath)  # このスクリプトのあるディレクトリのフルパスを取得。
	name = source.getImplementationName().split(".")[-1]
	global C
	filename = "".join((TIMESTAMP, "_", str(C), "{}_{}".format(name, filename), ".log"))
	C += 1
	with open(os.path.join(dirpath, filename), "w") as f:
		f.write(txt)