#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import Rectangle
from com.sun.star.awt import XMenuListener
from com.sun.star.awt.MenuItemStyle import AUTOCHECK, RADIOCHECK, CHECKABLE
from com.sun.star.awt.PopupMenuDirection import EXECUTE_DEFAULT
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
    doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
    docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
    docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
    toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。  
    dialog, addControl = dialogCreator(ctx, smgr, {"PositionX": 102, "PositionY": 41, "Width": 200, "Height": 140, "Title": "Menu-Dialog", "Name": "Dialog1", "Step": 1, "TabIndex": 0, "Moveable": True})
    addControl("FixedText", {"Name": "Headerlabel", "PositionX": 6, "PositionY": 6, "Width": 200, "Height": 8, "Label": "This code-sample demonstrates the creation of a popup-menu."})
    addControl("FixedText", {"Name": "Popup", "PositionX": 50, "PositionY": 50, "Width": 100, "Height": 8, "Label": "Right-click here"}, {"addMouseListener": MouseListener(ctx, smgr, dialog)})
    dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。
    # ノンモダルダイアログにするとき。
#     showModelessly(ctx, smgr, docframe, dialog)  
    # モダルダイアログにする。フレームに追加するとエラーになる。
    dialog.execute()  
    dialog.dispose()        
class MouseListener(unohelper.Base, XMouseListener):  # Editコントロールではうまく動かない。    
    def __init__(self, ctx, smgr, dialog):
        createMenu = menuCreator(ctx, smgr)
        menulistener = MenuListener(dialog)
        items = ("First Entry", CHECKABLE+AUTOCHECK, {"checkItem": True}),\
                ("First Radio Entry", RADIOCHECK+AUTOCHECK, {"enableItem": False}),\
                ("Second Radio Entry", RADIOCHECK+AUTOCHECK),\
                ("Third Radio Entry", RADIOCHECK+AUTOCHECK, {"checkItem": True}),\
                (),\
                ("Fifth Entry", CHECKABLE+AUTOCHECK),\
                ("Fourth Entry", CHECKABLE+AUTOCHECK, {"checkItem": True}),\
                ("Sixth Entry", 0),\
                ("~Close", 0, {"setCommand": "close"})
        self.popupmenu =  createMenu("PopupMenu", items, {"addMenuListener": menulistener})   
        items = ("First Entry", CHECKABLE+AUTOCHECK, {"checkItem": True}),\
                ("Second Entry", 0)
        popupmenu =  createMenu("PopupMenu", items, {"addMenuListener": menulistener}) 
        self.popupmenu.setPopupMenu (8, popupmenu)
#     @enableRemoteDebugging
    def mousePressed(self, mouseevent):
        control, dummy_controlmodel, name = eventSource(mouseevent)
        if name == "Popup":
            if mouseevent.PopupTrigger:
                pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)
                self.popupmenu.execute(control.getPeer(), pos, EXECUTE_DEFAULT)
    def mouseReleased(self, mouseevent):
        pass
    def mouseEntered(self, mouseevent):
        pass
    def mouseExited(self, mouseevent):
        pass
    def disposing(self, eventobject):
        pass
class MenuListener(unohelper.Base, XMenuListener):
    def __init__(self, dialog):
        self.dialog = dialog
    def itemHighlighted(self, menuevent):
        pass
#     @enableRemoteDebugging
    def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。
        cmd = menuevent.Source.getCommand(menuevent.MenuId)
        if cmd == "close":
            doc = XSCRIPTCONTEXT.getDocument()
            if doc:  # ドキュメントが取得できた時はモダルダイアログと判断する(汎用性は未確認)。
                self.dialog.endExecute()  # ウィンドウを閉じる。モダルダイアログではこれだけで閉じる。 モードレスダイアログは閉じない。
            else:
                self.dialog.dispose()  # モードレスダイアログではこれでウィンドウが閉じる(本来はフレームをdispose()すべき?)。モダルダイアログではdispose()ではドキュメントまでも閉じてしまう。
    def itemActivated(self, menuevent):
        pass
    def itemDeactivated(self, menuevent):
        pass   
    def disposing(self, eventobject):
        pass     
def menuCreator(ctx, smgr):
    def createMenu(menutype, items, attr=None):   
        if attr is None:
            attr = {}
        menu = smgr.createInstanceWithContext("com.sun.star.awt.{}".format(menutype), ctx) 
        for i, item in enumerate(items, start=1):
            if item:
                if len(item) > 2:
                    item = list(item)
                    attr[i] = item.pop()
                menu.insertItem(i, *item, i-1)  # ItemId, Text, ItemSytle, ItemPos。
            else:
                menu.insertSeparator(i)  # ItemPos    
        if attr:
            for key, val in attr.items():
                if isinstance(val, dict):
                    for method, arg in val.items():
                        if method in ("checkItem", "enableItem", "setCommand", "setHelpCommand", "setHelpText", "setTipHelpText"):
                            getattr(menu, method)(key, arg)
                else:
                    getattr(menu, key)(val)    
        return menu
    return createMenu    
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
        labels = None
        if controltype == "Roadmap":  # Roadmapコントロールのアイテム名の辞書を取得する。
            if "Items" in props:
                labels = props.pop("Items")
        if "PosSize" in props:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
            control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
            control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
            controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
            control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
            dialog.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
        else:  # Map AppFont (ma)のときはダイアログモデルにモデルを追加しないと正しくピクセルに変換されない。
            controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
            dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
        if labels is not None:  # コントロールに追加されたモデルでないとRoadmapアイテムは追加できない。
            i = 0
            for label in labels:  # 各Roadmapアイテムのラベルについて
                item = controlmodel.createInstance()
                item.setPropertyValues(("Label", "Enabled"), (label, True))     # IDは最小の自然数が自動追加されるので設定不要。
                controlmodel.insertByIndex(i, item)        
                i += 1
        if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
            control = dialog.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
            for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
                if val is None:
                    getattr(control, key)()
                else:
                    getattr(control, key)(val)
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
                return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')
#                 return self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
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