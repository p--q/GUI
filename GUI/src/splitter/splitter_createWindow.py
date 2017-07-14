#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.awt import XMouseMotionListener
from com.sun.star.awt.MouseButton import LEFT
from com.sun.star.awt.PosSize import POSSIZE, Y, HEIGHT
from com.sun.star.awt.WindowAttribute import  CLOSEABLE, SHOW, MOVEABLE, BORDER
from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.awt import Rectangle
from com.sun.star.awt import WindowDescriptor

def macro():
    ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
    smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
    doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。 
    docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
    docwindow = docframe.getContainerWindow()
    toolkit = docwindow.getToolkit()
    subwindow = createWindow(toolkit, docwindow, "modaldialog", BORDER + MOVEABLE + CLOSEABLE, 100, 100, 300, 200)
    controlcontainer = createControl(smgr, ctx, "Container", 0, 0, 300, 200, (), ())
    controlcontainer.setPosSize(0, 0, 300, 200, POSSIZE)
    controlcontainer.createPeer(toolkit, subwindow)
    edit1 = createControl(smgr, ctx, "Edit", 0, 0, 300, 99, ("AutoVScroll", "MultiLine",), (True, True,))
    edit2 = createControl(smgr, ctx, "Edit", 0, 105, 300, 96, ("AutoVScroll", "MultiLine",), (True, True,))
    controlcontainer.addControl("edit1",edit1)
    controlcontainer.addControl("edit2",edit2)
    spl = createWindow(toolkit, controlcontainer.getPeer(), "splitter", SHOW + BORDER, 0, 100, 300, 5)  # コントロールコンテナの上にさらにウィンドウを作成する。
    spl.setProperty("BackgroundColor", 0xEEEEEE)
    spl.setProperty("Border",1)
    spl.addMouseMotionListener(mouse_motion(controlcontainer))
    # モードレスダイアグにするｔき
    createFrame = frameCreator(ctx, smgr, docframe)
    createFrame("newFrame", subwindow)
    subwindow.setVisible(True)
    # モダルダイアログにするとき
#     subwindow.execute()
#     spl.dispose()
#     subwindow.dispose()
class mouse_motion(unohelper.Base, XMouseMotionListener):
    def __init__(self,controlcontainer):
        self.controlcontainer = controlcontainer
        self.edit1 = self.controlcontainer.getControl("edit1")
        self.edit2 = self.controlcontainer.getControl("edit2")
    def disposing(self,ev):
        pass
    def mouseMoved(self,ev):
        pass
    def mouseDragged(self,ev):
        if ev.Buttons ==  LEFT:
            split_pos = ev.Source.getPosSize()
            if 5 < split_pos.Y + ev.Y < 200:
                ev.Source.setPosSize(0, split_pos.Y + ev.Y, 0, 0, Y)
                edit1_pos = self.edit1.getPosSize()
                edit2_pos = self.edit2.getPosSize()
                self.edit1.setPosSize(0, 0, 0, edit1_pos.Height + ev.Y, HEIGHT)
                self.edit2.setPosSize(0, edit2_pos.Y + ev.Y, 0,edit2_pos.Height - ev.Y, Y + HEIGHT)
def createWindow(toolkit, parent, service, attr, nX, nY, nWidth, nHeight):
    aRect = Rectangle(X=nX, Y=nY, Width=nWidth, Height=nHeight)
    d = WindowDescriptor(Type=SIMPLE, WindowServiceName=service, ParentIndex=-1, Bounds=aRect, Parent=parent, WindowAttributes=attr)
    return toolkit.createWindow(d)
# create control by its type with property values
def createControl(smgr, ctx, ctype, x, y, width, height, names, values):
    ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % ctype, ctx)
    ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % ctype, ctx)
    ctrl_model.setPropertyValues(names, values)
    ctrl.setModel(ctrl_model)
    ctrl.setPosSize(x, y, width, height, POSSIZE)
    return ctrl
def frameCreator(ctx, smgr, parentframe): # 新しいフレームを追加する関数を返す。親フレームを渡す。   
    def createFrame(framename, containerwindow):  # 新しいフレーム名、そのコンテナウィンドウにするウィンドウを渡す。
        frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しく作成したウィンドウを入れるためのフレームを作成。
        frame.initialize(containerwindow)  # フレームにウィンドウを入れる。    
        frame.setName(framename)  # フレーム名を設定。
        parentframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。 
        return frame        
    return createFrame       
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。


if __name__ == "__main__":  # オートメーションで実行するとき
    import officehelper
    import traceback
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
                print("Could not establish a connection with a running office.")
                sys.exit()
            print("Connected to a running office ...")
            smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
            print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
            try:
                return func(ctx, smgr)  # 引数の関数の実行。
            except:
                traceback.print_exc()
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