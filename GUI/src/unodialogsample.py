#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
import unohelper
from com.sun.star.awt import XMouseListener
from com.sun.star.awt.SystemPointer import REFHAND
from com.sun.star.awt import XTextListener
from com.sun.star.awt import XFocusListener
from com.sun.star.awt.MessageBoxType import ERRORBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK


def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
    dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    props = {"PositionX": 102, "PositionY": 41, "Width": 380, "Height": 380, "Title": "LibreOffice", "Name": "MyTestDialog", "Step": 0, "TabIndex": 0, "Moveable": True}
    dialogmodel.setPropertyValues(tuple(props.keys()), tuple(props.values()))
    dialogview = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx) 
    controls = ControlCreator(dialogmodel, dialogview)
    props = {"PositionX": 106, "PositionY": 6, "Width": 300, "Height": 8, "Label": "This code-sample demonstrates how to create various controls in a dialog"}
    controls.add("FixedText", props, name="Headerlabel")
    props = {"PositionX": 106, "PositionY": 18, "Width": 100, "Height": 8, "Label": "My ~Label", "Step": 0}
    controls.add("FixedText", props, listeners={"MouseListener": MouseListener(ctx, smgr)})
    props = {"PositionX": 106, "PositionY": 30, "Width": 60, "Height": 12, "PrependCurrencySymbol": True, "CurrencySymbol": "$", "Value": 2.93}
    controls.add("CurrencyField", props)
    props = {"PositionX": 106, "PositionY": 44, "Width": 250, "Height": 8, "ProgressValueMin": 0, "ProgressValueMax": 100}
    controls.add("ProgressBar", props)
    props = {"PositionX": 106, "PositionY": 58, "Width": 100, "Height": 8, "Orientation": 0, "Label": "My FixedLine"}
    controls.add("FixedLine", props)   
    props = {"PositionX": 106, "PositionY": 72, "Width": 60, "Height": 12, "Text": "MyText", "EchoChar": ord("*")}
    controls.add("Edit", props, listeners={"TextListener": TextListener(), "FocusListener": FocusListener()})  
    props = {"PositionX": 106, "PositionY": 58, "Width": 100, "Height": 8, "Orientation": 0, "Label": "My FixedLine"}
    controls.add("FixedLine", props)      
    
    
    
# (106, 96, 50, new Time(0,(short)0,(short)0,(short)10,false), new Time((short)0,(short)0,(short)0,(short)0,false), new Time((short)0,(short)0,(short)0,(short)17,false))


    dialogview.setVisible(True)
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    dialogview.createPeer(toolkit, None)
#     peer = unodialog.getPeer()
#     activateProgressBar(dialogview)
    dialogview.execute()
    
    
def activateProgressBar(dialogview):
    progressbarview = dialogview.getControl("ProgressBar1")
    progressbarmodel = progressbarview.getModel()
    progressbarmodel.BackgroundColor = 0x0000ff
    import time
    for i in range(0, 101, 20):
        progressbarmodel.ProgressValue = i
        time.sleep(1)
class ControlCreator:
    def __init__(self, dialogmodel, dialogview):
        dialogview.setModel(dialogmodel)   
        self.model = dialogmodel
        self.view = dialogview
    def add(self, controltype, props, *, name=None, listeners=None):#             peer = ev.Source.getPeer()
        name = name or self.generateSequentialName(controltype)
        if not "Name" in props:
            props["Name"] = name  
        controlmodel = self.model.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))
        controlmodel.setPropertyValues(tuple(props.keys()), tuple(props.values()))
        self.model.insertByName(name, controlmodel)
        if listeners is not None:
            controlview = self.view.getControl(name)
            for key, val in listeners.items():
                getattr(controlview, "add{}".format(key))(val)
    def generateSequentialName(self, controltype):
        i = 1
        flg = True
        while flg:
            name = "{}{}".format(controltype, i)
            flg = self.model.hasByName(name)
            i += 1
        return name
class MouseListener(unohelper.Base, XMouseListener):    
    def __init__(self, ctx, smgr):
        pointer = smgr.createInstanceWithContext("com.sun.star.awt.Pointer", ctx)
        pointer.setType(REFHAND)
        self.pointer = pointer 
    def mousePressed(self, ev):
        pass
    def mouseReleased(self, ev):
        pass
    def mouseEntered(self, ev):
        ctrl = ev.Source
        ctrl.getPeer().setPointer(self.pointer)
    def mouseExited(self, ev):
        pass
    def disposing(self, src):
        pass
class TextListener(unohelper.Base, XTextListener):
#     def __init__(self, ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
#         self.desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)  # コンポーネントコンテクストからデスクトップをインスタンス化。
# 
#         self.ctx = ctx
    
    def textChanged(self, ev):
#         import pydevd; pydevd.settrace()
        controlmodel = ev.Source.getModel()
        name = controlmodel.getPropertyValue("Name")
        if name=="Edit1":
            txt = controlmodel.getPropertyValue("Text")
            peer = ev.Source.getPeer()
            msgbox = peer.getToolkit().createMessageBox(peer, ERRORBOX, BUTTONS_OK, "TextListener", "EditModel has changed to '{}'".format(txt))  # ピアオブジェクトからツールキットを取得して、peerを親ウィンドウにしてメッセージボックスを作成。
            msgbox.execute()  # メッセージボックスを表示。
            msgbox.dispose()  # メッセージボックスを破棄。


#         from msgbox import MsgBox
#         s = txt
#         myBox = MsgBox(self.ctx)
#         myBox.addButton("OK")
#         myBox.renderFromBoxSize(100)  # メッセージボックスの幅を文字列の長さから算出。
#         myBox.show(s, 0, "Content of Edit1 control")



    def disposing(self, src):
        pass    
class FocusListener(unohelper.Base, XFocusListener):
    def focusGained(self, ev):
        pass
    def focusLost(self, ev):
        pass

    def disposing(self, src):
        pass  

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
        if not smgr:
            print( "ERROR: no service manager" )
            sys.exit()
        print("Using remote servicemanager\n") 
        try:
            func(ctx, smgr)  # 引数の関数の実行。
        except:
            traceback.print_exc()
#         soffice.binの終了処理。これをしないとLibreOfficeを起動できなくなる。
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        prop = PropertyValue(Name="Hidden", Value=True)
        desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))  # バックグラウンドでWriterのドキュメントを開く。
        terminated = desktop.terminate()  # LibreOfficeをデスクトップに展開していない時はエラーになる。
        if terminated:
            print("\nThe Office has been terminated.")  # 未保存のドキュメントがないとき。
        else:
            print("\nThe Office is still running. Someone else prevents termination.")  # 未保存のドキュメントがあってキャンセルボタンが押された時。
    return wrapper
if __name__ == "__main__":
    main = connectOffice(main)
    main()