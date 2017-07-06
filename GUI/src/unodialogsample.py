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
from com.sun.star.awt.FocusChangeReason import TAB
from com.sun.star.util import Time
from com.sun.star.util import Date
from com.sun.star.awt import XSpinListener
from com.sun.star.awt import XItemListener


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
    time = Time(Hours=10, Minutes=0, Seconds=0, NanoSeconds=0, IsUTC=False)
    timemin = Time(Hours=0, Minutes=0, Seconds=0, NanoSeconds=0, IsUTC=False) 
    timemax = Time(Hours=17, Minutes=0, Seconds=0, NanoSeconds=0, IsUTC=False) 
    props = {"PositionX": 106, "PositionY": 96, "Width": 50, "Height": 12, "Spin": True, "TimeFormat": 5, "Time": time, "TimeMin": timemin, "TimeMax": timemax}
    controls.add("TimeField", props)      
    date = Date(Year=2017, Month=7, Day=4)
    datemin = Date(Year=2017, Month=6, Day=16)
    datemax = Date(Year=2017, Month=8, Day=15)
    props = {"PositionX": 166, "PositionY": 96, "Width": 50, "Height": 12, "Dropdown": True, "DateFormat": 9, "DateMin": date, "DateMax": datemin, "Date": datemax}
    controls.add("DateField", props)     
    props = {"PositionX": 102, "PositionY": 124, "Width": 100, "Height": 70, "Label": "~My GroupBox"}
    controls.add("GroupBox", props)   
    props = {"PositionX": 106, "PositionY": 136, "Width": 50, "Height": 12, "LiteralMask": "__.05.2007", "EditMask": "NNLLLLLLLL", "StrictFormat": True}
    controls.add("PatternField", props)   
    props = {"PositionX": 106, "PositionY": 152, "Width": 50, "Height": 12, "Spin": True, "StrictFormat": True, "ValueMin": 0.0, "ValueMax": 1000.0, "Value": 500.0, "ValueStep": 100.0, "ShowThousandsSeparator": True, "DecimalAccuracy": 1}
    controls.add("NumericField", props)  
    props = {"PositionX": 106, "PositionY": 168, "Width": 150, "Height": 8, "Label": "~Include page number", "TriState": True, "State": 1}
    controls.add("CheckBox", props, listeners={"ItemListener": ItemListener(dialogview)})  
    props = {"PositionX": 130, "PositionY": 200, "Width": 150, "Height": 8, "Label": "~First Option", "State": 1, "TabIndex": 51}
    controls.add("RadioButton", props)     
    props = {"PositionX": 130, "PositionY": 214, "Width": 150, "Height": 8, "Label": "~Second Option", "TabIndex": 50}
    controls.add("RadioButton", props)      
    props = {"PositionX": 106, "PositionY": 230, "Width": 50, "Height": 30, "Dropdown": False, "Step": 0, "MultiSelection": True, "StringItemList": ("First Item", "Second Item", "ThreeItem"), "SelectedItems": (0, 2)}
    controls.add("ListBox", props)     
    props = {"PositionX": 160, "PositionY": 230, "Width": 50, "Height": 12, "Dropdown": True, "MaxTextLen": 10, "ReadOnly": False, "StringItemList": ("First Entry", "Second Entry", "Third Entry", "Fourth Entry")}   
    controls.add("ComboBox", props)  







#     combobox1 = dialogview.getControl("ComboBox1")
#     print(combobox1.Text)

    dialogview.setVisible(True)
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    dialogview.createPeer(toolkit, None)
#     peer = unodialog.getPeer()
#     activateProgressBar(dialogview)
    dialogview.execute()
    dialogview.dispose()
    
def activateProgressBar(dialogview):
    progressbarcontroller = dialogview.getControl("ProgressBar1")
    progressbarmodel = progressbarcontroller.getModel()
    progressbarmodel.BackgroundColor = 0x0000ff
    import time
    for i in range(0, 101, 20):
        progressbarmodel.ProgressValue = i
        time.sleep(1)
class ControlCreator:
    def __init__(self, dialogmodel, dialogview):
        dialogview.setModel(dialogmodel)   
        self.model = dialogmodel
        self.controller = dialogview
    def add(self, controltype, props, *, name=None, listeners=None):
        name = name or self.generateSequentialName(controltype)
        if not "Name" in props:
            props["Name"] = name  
        controlmodel = self.model.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))
        [setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。
#         controlmodel.setPropertyValues(tuple(props.keys()), tuple(props.values()))  # タプルが値の時はすべて[]anyと判断されるのでエラーが出る。
        self.model.insertByName(name, controlmodel)
        if listeners is not None:
            controlview = self.controller.getControl(name)
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
    def mousePressed(self, mouseevent):
        pass
    def mouseReleased(self, mouseevent):
        pass
    def mouseEntered(self, mouseevent):
        ctrl = mouseevent.Source
        ctrl.getPeer().setPointer(self.pointer)
    def mouseExited(self, mouseevent):
        pass
    def disposing(self, eventobject):
        pass
class TextListener(unohelper.Base, XTextListener):
    def textChanged(self, textevent):
        controlmodel = textevent.Source.getModel()
        name = controlmodel.getPropertyValue("Name")
        if name=="Edit1":
            txt = controlmodel.getPropertyValue("Text")
            peer = textevent.Source.getPeer()
            msgbox = peer.getToolkit().createMessageBox(peer, ERRORBOX, BUTTONS_OK, "TextListener", "EditModel has changed to '{}'".format(txt))  # ピアオブジェクトからツールキットを取得して、peerを親ウィンドウにしてメッセージボックスを作成。
            msgbox.execute()  # メッセージボックスを表示。
            msgbox.dispose()  # メッセージボックスを破棄。
    def disposing(self, eventobject):
        pass    
class FocusListener(unohelper.Base, XFocusListener):
    def focusGained(self, focusevent):
        pass
    def focusLost(self, focusevent):
        focusflags = focusevent.FocusFlags
        focuschangereason = focusflags & TAB
        if focuschangereason==TAB:
            controlwindow = focusevent.NextFocus
    def disposing(self, eventobject):
        pass  
class SpinListener(unohelper.Base, XSpinListener):
    def up(self, spinevent):
        controlmodel = spinevent.Source.getModel()
        
    def down(self, spinevent):
        pass
    def first(self, spinevent):
        pass
    def last(self, spinevent):
        pass
    def disposing(self, eventobject):
        pass  
class ItemListener(unohelper.Base, XItemListener): 
    def __init__(self, dialogview):
        self.view = dialogview
    def itemStateChanged(self, itemevent):
        checkbox = itemevent.Source
        controlview = self.view.getControl("CommandButton1")
        controlmodel = controlview.getModel()
        state = checkbox.getState()
        bdoenable = True
        if state==0 or state==2:
            bdoenable = False
        controlmodel.setPropertyValue("Enabled", bdoenable)
    def disposing(self, eventobject):
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
        print("Using {} {}".format(*getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
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
def getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
    cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
    node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
    ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
    return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
if __name__ == "__main__":
    main = connectOffice(main)
    main()