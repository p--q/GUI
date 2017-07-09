#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
import unohelper
from com.sun.star.awt import XMouseMotionListener
from com.sun.star.awt.MouseButton import LEFT
from com.sun.star.awt.PosSize import POSSIZE, SIZE, Y, HEIGHT
from com.sun.star.awt.WindowAttribute import  CLOSEABLE, SHOW, MOVEABLE, BORDER
import uno
from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.beans import NamedValue
from com.sun.star.awt import Rectangle


def main(ctx, smgr):
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    frame = desktop.getActiveFrame()
    parent = frame.getContainerWindow()
    toolkit = parent.getToolkit()
    subwin = createWindow(toolkit, parent, "modaldialog", SHOW + BORDER + MOVEABLE + CLOSEABLE, 100, 100, 300, 200)
    
#     taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
#     args = NamedValue("PosSize", Rectangle(100, 100, 300, 200)), 
#     frame = taskcreator.createInstanceWithArguments(args)
#     subwin = frame.getContainerWindow()
#     toolkit = subwin.getToolkit()

    
    cont = createControl(smgr, ctx, "Container", 0, 0, 300, 200, (), ())
    cont.setPosSize(0, 0, 300, 200, POSSIZE)
    cont.createPeer(toolkit, subwin)
    
#     frame.setComponent(cont, None)
    
    edit1 = createControl(smgr, ctx, "Edit", 0, 0, 300, 99, ("AutoVScroll", "MultiLine",), (True, True,))
    edit2 = createControl(smgr, ctx, "Edit", 0, 105, 300, 96, ("AutoVScroll", "MultiLine",), (True, True,))
    cont.addControl("edit1",edit1)
    cont.addControl("edit2",edit2)
    spl = createWindow(toolkit, cont.getPeer(), "splitter", SHOW + BORDER, 0, 100, 300, 5)
    spl.setProperty("BackgroundColor", 0xEEEEEE)
    spl.setProperty("Border",1)
    spl.addMouseMotionListener(mouse_motion(cont))
    #inspect(ctx,spl)
    subwin.execute()

#     subwin.setVisible(True)
    
    
    spl.dispose()
    subwin.dispose()
class mouse_motion(unohelper.Base, XMouseMotionListener):
    def __init__(self,cont):
        self.cont = cont
        self.edit1 = self.cont.getControl("edit1")
        self.edit2 = self.cont.getControl("edit2")
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
    d = uno.createUnoStruct("com.sun.star.awt.WindowDescriptor")
    d.Type = SIMPLE
    d.WindowServiceName = service
    d.ParentIndex = -1
    d.Bounds = createRect(nX, nY, nWidth, nHeight)
    d.Parent = parent
    d.WindowAttributes = attr
    return toolkit.createWindow(d)
def createRect(x,y,width,height):
    aRect = uno.createUnoStruct("com.sun.star.awt.Rectangle")
    aRect.X = x
    aRect.Y = y
    aRect.Width = width
    aRect.Height = height
    return aRect
# create control by its type with property values
def createControl(smgr, ctx, ctype, x, y, width, height, names, values):
    ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % ctype, ctx)
    ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % ctype, ctx)
    ctrl_model.setPropertyValues(names, values)
    ctrl.setModel(ctrl_model)
    ctrl.setPosSize(x, y, width, height, POSSIZE)
    return ctrl


# funcの前後でOffice接続の処理
def connectOffice(func):
    @wraps(func)
    def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
        ctx = None
        try:
            ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
        except:
            pass
        if not ctx:
            print("Could not establish a connection with a running office.")
            sys.exit()
        print("Connected to a running office ...")
        smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
        print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
        try:
            func(ctx, smgr)  # 引数の関数の実行。
        except:
            traceback.print_exc()
#         _terminateOffice(ctx, smgr) # soffice.binの終了処理。これをしないとLibreOfficeを起動できなくなる。    
    def _terminateOffice(ctx, smgr):  # soffice.binの終了処理。
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        prop = PropertyValue(Name="Hidden", Value=True)
        desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))  # バックグラウンドでWriterのドキュメントを開く。
        terminated = desktop.terminate()  # LibreOfficeをデスクトップに展開していない時はエラーになる。
        if terminated:
            print("\nThe Office has been terminated.")  # 未保存のドキュメントがないとき。
        else:
            print("\nThe Office is still running. Someone else prevents termination.")  # 未保存のドキュメントがあってキャンセルボタンが押された時。
    def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
        cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
        node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
        ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
        return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
    return wrapper
if __name__ == "__main__":
    main = connectOffice(main)
    main()