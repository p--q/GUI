#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
import uno
from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.style.VerticalAlignment import BOTTOM
import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.awt.MessageBoxType import INFOBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.awt.WindowAttribute import  CLOSEABLE, SHOW, MOVEABLE, BORDER
from com.sun.star.awt.PushButtonType import OK 


def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
  
#     prop = PropertyValue(Name="Hidden", Value=True)
#     desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))  # バックグラウンドでWriterのドキュメントを開く。
      
    doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
    
    
# def macro():
#     ctx = XSCRIPTCONTEXT.getComponentContext()
#     smgr = ctx.getServiceManager()
#     doc = XSCRIPTCONTEXT.getDocument()   

    
    
    parentframe = doc.getCurrentController().getFrame()
    peer = parentframe.getContainerWindow()
    toolkit = peer.getToolkit()
    window =  createWindow(toolkit, peer, "dialog", SHOW + BORDER + MOVEABLE + CLOSEABLE, 150, 150, 200, 200)
    frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)
    frame.initialize(window)
#     frame.setCreator(parentframe)
    frame.setName("NewFrame")
#     frame.setTitle("New Frame")
#     parentframe.getFrames().append(frame)
    window.setVisible(True)
    controlcontainer = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)
    controlcontainermodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)
    controlcontainermodel.setPropertyValue("BackgroundColor", -1)
    controlcontainer.setModel(controlcontainermodel)
    controlcontainer.createPeer(toolkit, window)
    controlcontainer.setPosSize(0, 0, 200, 200, POSSIZE)
    frame.setComponent(controlcontainer, None)
    label = createControl(smgr, ctx, "FixedText", 10, 0, 180, 30, ("Label", "VerticalAlign"), ("Label1", BOTTOM))
    edit = createControl(smgr, ctx, "Edit", 10, 40, 180, 30, (), ())
    btn = createControl(smgr, ctx, "Button", 110, 130, 80, 35, ("DefaultButton", "Label"), (True, "btn"))

    controlcontainer.addControl("label", label)
    controlcontainer.addControl("btn", btn)
    controlcontainer.addControl("edit", edit)
    
    edit.setFocus()
    btn.setActionCommand("btn") 
    btn.addActionListener(BtnListener(controlcontainer))
     
     
        
#     window.setVisible(True)  # モードレスダイアログ。オートメーションではボタンが起動しない。

#     msgbox = toolkit.createMessageBox(window, INFOBOX, BUTTONS_OK, "Text Field", "{}".format("ctrl.getText()"))  # ピアオブジェクトからツールキットを取得して、peerを親ウィンドウにしてメッセージボックスを作成。
#     msgbox.execute() 

#     window.execute()  # モダルダイアログ。
    
#     window.dispose()
    
class BtnListener(unohelper.Base, XActionListener):
    def __init__(self, controlcontainer):
        self.controlcontainer = controlcontainer
    def actionPerformed(self, actionevent):
#         import pydevd; pydevd.settrace()
#         cmd = actionevent.ActionCommand
#         ctrl = actionevent.Source
        
        editcontrol = self.controlcontainer.getControl("edit")
        editcontrol.setText("By Button Click")
        
#         toolkit = self.peer.getToolkit()
#         if cmd == "btn":
#             
#             ctrl.setText("By Button Click")
#             
#             msgbox = toolkit.createMessageBox(self.peer, INFOBOX, BUTTONS_OK, "Text Field", "{}".format(ctrl.getText()))  # ピアオブジェクトからツールキットを取得して、peerを親ウィンドウにしてメッセージボックスを作成。
#             msgbox.execute()  # メッセージボックスを表示。
#             msgbox.dispose()  # メッセージボックスを破棄。
    def disposing(self, eventobject):
        pass





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
    ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(ctype), ctx)
    ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}Model".format(ctype), ctx)
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