#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.awt import Rectangle
from com.sun.star.awt.MenuItemStyle import AUTOCHECK, RADIOCHECK, CHECKABLE
from com.sun.star.awt import XMenuListener
from com.sun.star.util import XCloseListener
from com.sun.star.beans import NamedValue
from com.sun.star.awt.PosSize import POSSIZE
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
    if __name__ == "__main__":  # オートメーションのときはデバッグサーバーに接続しない。
        return func
    def wrapper(*args, **kwargs):
        import time
        import pydevd
        doc = XSCRIPTCONTEXT.getDocument()
        if doc:  # ドキュメントが取得できた時
            indicator = doc.getCurrentController().getFrame().createStatusIndicator()  # フレームからステータスバーを取得する。
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
def macro():  # オートメーションではリスナーが呼ばれない、閉じるボタンでウィンドウを閉じるとLibreOfficeがクラッシュする。
    ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
    smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
    doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
    docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
    docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
    toolkit = docwindow.getToolkit()  # ツールキットを取得。
    taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
    args = NamedValue("PosSize", Rectangle(100, 100, 500, 500)), NamedValue("FrameName", "NewFrame")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
    frame = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
    window = frame.getContainerWindow()  # 新しいコンテナウィンドウを新しいフレームから取得。
    frame.setTitle("MenuBar Example")  # フレームのタイトルを設定。
    docframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。
    createMenu = menuCreator(ctx, smgr)
    items = ("~First MenuBar Item", 0),\
            ("~Second MenuBar Item", 0)    
    menubar = createMenu("MenuBar", items)
    window.setMenuBar(menubar)
    menulistener = MenuListener(window, menubar)
    items = ("First Entry", CHECKABLE+AUTOCHECK, {"checkItem": True}),\
            ("First Radio Entry", RADIOCHECK+AUTOCHECK, {"enableItem": False}),\
            ("Second Radio Entry", RADIOCHECK+AUTOCHECK),\
            ("Third Radio Entry", RADIOCHECK+AUTOCHECK, {"checkItem": True}),\
            (),\
            ("Fifth Entry", CHECKABLE+AUTOCHECK),\
            ("Fourth Entry", CHECKABLE+AUTOCHECK, {"checkItem": True}),\
            ("Sixth Entry", 0),\
            ("~Close", 0, {"setCommand": "close"})
    popupmenu =  createMenu("PopupMenu", items, {"addMenuListener": menulistener})   
    menubar.setPopupMenu(1, popupmenu)  
    items = ("First Entry", CHECKABLE+AUTOCHECK, {"checkItem": True}),\
            ("Second Entry", 0)
    popupmenu =  createMenu("PopupMenu", items, {"addMenuListener": menulistener})       
    menubar.setPopupMenu(2, popupmenu)  
    controlcontainer, addControl = controlcontainerCreator(ctx, smgr, {"PositionX": 0, "PositionY": 0, "Width": 500, "Height": 500, "PosSize": POSSIZE})
    addControl("FixedText", {"PositionX": 20, "PositionY": 20, "Width": 460, "Height": 460, "PosSize": POSSIZE, "Label": "\n".join(getStatus(menubar)), "NoLabel": True, "MultiLine": True})
    menulistener.control = controlcontainer.getControl("FixedText1")
    controlcontainer.createPeer(toolkit, window) 
    controlcontainer.setVisible(True)
    window.setVisible(True)
def getStatus(menubar):
    t = "    "
    txts = []
    c = menubar.getItemCount()
    for i in range(c):
        flg = True
        popup = menubar.getPopupMenu(i+1)
        pc = popup.getItemCount()
        for j in range(pc):
            if popup.isItemChecked(j+1):
                if flg:
                    txts.append(menubar.getItemText(i+1).replace("~", ""))
                    flg = False
                txts.append("{}{} is checked.".format(t, popup.getItemText(j+1).replace("~", "")))
        txts.append("")
    return txts
class MenuListener(unohelper.Base, XMenuListener):
    def __init__(self, window, menubar):
        self.window = window
        self.menubar = menubar
        self.control = None
    def itemHighlighted(self, menuevent):
        pass
#     @enableRemoteDebugging
    def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。
        cmd = menuevent.Source.getCommand(menuevent.MenuId)
        if cmd == "close":
            self.window.dispose()  # ウィンドウを閉じる。      
        else:  # com::sun::star::awt::MenuItemStyleを知る方法がみつからない。uncheckedを知る方法もない。
            self.control.getModel().setPropertyValue("Label", "\n".join(getStatus(self.menubar)))
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
def controlcontainerCreator(ctx, smgr, containerprops):  # コントロールコンテナと、それにコントロールを追加する関数を返す。まずコントロールコンテナモデルのプロパティを取得。
    container = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)  # コントロールコンテナの生成。
    if "PosSize" in containerprops:  # コントロールコンテナのコントロールはma単位は使えずピクセル単位で設定をする。
        container.setPosSize(containerprops.pop("PositionX"), containerprops.pop("PositionY"), containerprops.pop("Width"), containerprops.pop("Height"), containerprops.pop("PosSize"))
    containermodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)  # コンテナモデルの生成。
    containermodel.setPropertyValues(tuple(containerprops.keys()), tuple(containerprops.values()))  # コンテナモデルのプロパティを設定。
    container.setModel(containermodel)  # コンテナにコンテナモデルを設定。
    container.setVisible(False)  # 描画中のものを表示しない。
    def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
        labels = None
        if controltype == "Roadmap":  # Roadmapコントロールのアイテム名の辞書を取得する。
            if "Items" in props:
                labels = props.pop("Items")
        if "PosSize" in props:  # ピクセル単位でコントロールに設定をする。
            control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
            control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
            controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
            control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
            container.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
        else:  
            raise RuntimeError("You can not use ma unit in the controlmodel of a controlcontainermodel.")
        if labels is not None:  # コントロールに追加されたモデルでないとRoadmapアイテムは追加できない。
            i = 0
            for label in labels:  # 各Roadmapアイテムのラベルについて
                item = controlmodel.createInstance()
                item.setPropertyValues(("Label", "Enabled"), (label, True))     # IDは最小の自然数が自動追加されるので設定不要。
                controlmodel.insertByIndex(i, item)        
                i += 1
        if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
            control = container.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
            for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
                if val is None:
                    getattr(control, key)()
                else:
                    getattr(control, key)(val)
    def _createControlModel(controltype, props):  # コントロールモデルの生成。
        if not "Name" in props:
            props["Name"] = _generateSequentialName(controltype)  # Nameがpropsになければ通し番号名を生成。
        controlmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}Model".format(controltype), ctx)  # コントロールモデルを生成。    
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
            flg = container.getControl(name)  # 同名のコントロールの有無を判断。
            i += 1
        return name
    return container, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
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
                return self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
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