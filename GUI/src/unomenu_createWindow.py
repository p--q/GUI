#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.awt import WindowDescriptor
from com.sun.star.awt import Rectangle
from com.sun.star.awt.WindowClass import TOP
from com.sun.star.awt.WindowAttribute import  CLOSEABLE, SHOW, MOVEABLE, SIZEABLE
from com.sun.star.awt.MenuItemStyle import AUTOCHECK, CHECKABLE, RADIOCHECK
from com.sun.star.awt import XMenuListener
from com.sun.star.util import XCloseListener
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
def macro():  # オートメーションでは発火しないリスナーがある、閉じるボタンでウィンドウを閉じるとLibreOfficeがクラッシュする。
    ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
    smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
    doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
    docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
    docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
    toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。 
    window = createWindow(toolkit, CLOSEABLE+SHOW+MOVEABLE+SIZEABLE, {"PositionX": 100, "PositionY": 100, "Width": 500, "Height": 500, "Type": TOP})  
    window.setVisible(False)  # 描画中のウィンドウは表示しない。
    dummy_frame = addToFrames(ctx, smgr, docframe, window)  # 親フレームとウィンドウを渡す。新しいフレームのコンテナウィンドウにする。
    menubar = smgr.createInstanceWithContext("com.sun.star.awt.MenuBar", ctx)  # メニューバーをインスタンス化。
    menubar.insertItem(1, "~First MenuBar Item", 0, 0)  # ID(1から開始), ラベル、スタイル(0または定数com.sun.star.awt.MenuItemStyleの和)、位置(0から開始)
    menubar.insertItem(2, "~Second MenuBar Item", 0, 1)
    window.setMenuBar(menubar)  # メニューバーをウィンドウに追加。UnoControlDialogに追加しても目に見えない。
    items = ("First Entry", AUTOCHECK),\
            ("First Radio Entry", RADIOCHECK+AUTOCHECK),\
            ("Second Radio Entry", RADIOCHECK+AUTOCHECK),\
            ("Third Radio Entry", RADIOCHECK+AUTOCHECK),\
            (),\
            ("Fifth Entry", AUTOCHECK),\
            ("Fourth Entry", AUTOCHECK),\
            ("Sixth Entry", 0),\
            ("Close Dialog", 0)  # ポップアップメニューの項目。タプルのインデックスが位置に相当。IDはインデックス+1。
    popupmenu = createPopupMenu(ctx, smgr, items)  # ポップメニューを取得。
    popupmenu.enableItem(2, False)  # メニュー項目をIDで指定して、Falseでグレーアウト。
    popupmenu.checkItem(3, True)  # チェックできるメニュー項目をIDで指定して、Trueでチェックを付ける。
    popupmenu.addMenuListener(MenuListener(window))  # ポップアップメニューにリスナを付ける。
    menubar.setPopupMenu(1, popupmenu)  # メニューバーのIDを指定してポップアップメニューを追加。  
    window.setVisible(True)
class MenuListener(unohelper.Base, XMenuListener):
    def __init__(self, window):
        self.window = window
    def itemHighlighted(self, menuevent):
        pass
#     @enableRemoteDebugging
    def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。
        if menuevent.MenuId == 9:  # メニュー項目のIDを取得できる。
            self.window.dispose()  # ウィンドウを閉じる。
    def itemActivated(self, menuevent):
        pass
    def itemDeactivated(self, menuevent):
        pass   
    def disposing(self, eventobject):
        pass  
def createPopupMenu(ctx, smgr, items):  # ポップメニューを返す。itemsは項目のラベルとスタイルのタプルのタプル。   
    popupmenu = smgr.createInstanceWithContext("com.sun.star.awt.PopupMenu", ctx)  # ポップアップメニューをインスタンス化。
    for i, item in enumerate(items, start=1):  # 1から始まるiをIDにする。
        if item:  # ラベルとスタイルのタプルが取得できた時。
            popupmenu.insertItem(i, *item, i-1)  # ItemId, Text, ItemSytle, ItemPos。
        else:
            popupmenu.insertSeparator(i)  # ItemPos。セパレーターのときは位置を設定するだけ。    
    return popupmenu
def createWindow(toolkit, attr, props):  # ウィンドウタイトルは変更できない。attrはcom.sun.star.awt.WindowAttributeの和。propsはPositionX, PositionY, Width, Height, ParentIndex。
    aRect = Rectangle(X=props.pop("PositionX"), Y=props.pop("PositionY"), Width=props.pop("Width"), Height=props.pop("Height"))
    d = WindowDescriptor(Bounds=aRect, WindowAttributes=attr)
    for key, val in props.items():
        setattr(d, key, val)
    return toolkit.createWindow(d)  # ウィンドウピアを返す。
def addToFrames(ctx, smgr, parentframe, windowpeer):  # フレームに追加しないと閉じるボタンが使えない。
    frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しいフレームを生成。
    frame.initialize(windowpeer)  # フレームにコンテナウィンドウを入れる。    
    parentframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。 
    return frame  # フレームにリスナーをつけるときのためにフレームを返す。
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