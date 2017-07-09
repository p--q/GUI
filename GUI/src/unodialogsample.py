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
    dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログのモデルを取得。すべてのコントロールのモデルの入れ物。
    props = {"PositionX": 102, "PositionY": 41, "Width": 380, "Height": 380, "Title": "LibreOffice", "Name": "MyTestDialog", "Step": 0, "TabIndex": 0, "Moveable": True, "Tag": "ダイアログ"}
    dialogmodel.setPropertyValues(tuple(props.keys()), tuple(props.values()))
    addControlModel = controlmodelCreator(dialogmodel)
    addControlModel("FixedText", {"Name": "Headerlabel", "PositionX": 106, "PositionY": 6, "Width": 300, "Height": 8, "Label": "This code-sample demonstrates how to create various controls in a dialog"})
    addControlModel("FixedText", {"PositionX": 106, "PositionY": 18, "Width": 100, "Height": 8, "Label": "My ~Label", "Step": 0})
    addControlModel("CurrencyField", {"PositionX": 106, "PositionY": 30, "Width": 60, "Height": 12, "PrependCurrencySymbol": True, "CurrencySymbol": "$", "Value": 2.93})
    addControlModel("ProgressBar", {"PositionX": 106, "PositionY": 44, "Width": 250, "Height": 8, "ProgressValueMin": 0, "ProgressValueMax": 100})
    addControlModel("FixedLine", {"PositionX": 106, "PositionY": 58, "Width": 100, "Height": 8, "Orientation": 0, "Label": "My FixedLine"})   
    addControlModel("Edit", {"PositionX": 106, "PositionY": 72, "Width": 60, "Height": 12, "Text": "MyText", "EchoChar": ord("*")})  
    addControlModel("FixedLine", {"PositionX": 106, "PositionY": 58, "Width": 100, "Height": 8, "Orientation": 0, "Label": "My FixedLine"}) 
    time = Time(Hours=10, Minutes=0, Seconds=0, NanoSeconds=0, IsUTC=False)
    timemin = Time(Hours=0, Minutes=0, Seconds=0, NanoSeconds=0, IsUTC=False) 
    timemax = Time(Hours=17, Minutes=0, Seconds=0, NanoSeconds=0, IsUTC=False) 
    addControlModel("TimeField", {"PositionX": 106, "PositionY": 96, "Width": 50, "Height": 12, "Spin": True, "TimeFormat": 5, "Time": time, "TimeMin": timemin, "TimeMax": timemax})      
    date = Date(Year=2017, Month=7, Day=4)
    datemin = Date(Year=2017, Month=6, Day=16)
    datemax = Date(Year=2017, Month=8, Day=15)
    addControlModel("DateField", {"PositionX": 166, "PositionY": 96, "Width": 50, "Height": 12, "Dropdown": True, "DateFormat": 9, "DateMin": date, "DateMax": datemin, "Date": datemax})     
    addControlModel("GroupBox", {"PositionX": 102, "PositionY": 124, "Width": 100, "Height": 70, "Label": "~My GroupBox"})   
    addControlModel("PatternField", {"PositionX": 106, "PositionY": 136, "Width": 50, "Height": 12, "LiteralMask": "__.05.2007", "EditMask": "NNLLLLLLLL", "StrictFormat": True})   
    addControlModel("NumericField", {"PositionX": 106, "PositionY": 152, "Width": 50, "Height": 12, "Spin": True, "StrictFormat": True, "ValueMin": 0.0, "ValueMax": 1000.0, "Value": 500.0, "ValueStep": 100.0, "ShowThousandsSeparator": True, "DecimalAccuracy": 1})  
    addControlModel("CheckBox", {"PositionX": 106, "PositionY": 168, "Width": 150, "Height": 8, "Label": "~Include page number", "TriState": True, "State": 1})  
    addControlModel("RadioButton", {"PositionX": 130, "PositionY": 200, "Width": 150, "Height": 8, "Label": "~First Option", "State": 1, "TabIndex": 51})     
    addControlModel("RadioButton", {"PositionX": 130, "PositionY": 214, "Width": 150, "Height": 8, "Label": "~Second Option", "TabIndex": 50})      
    addControlModel("ListBox", {"PositionX": 106, "PositionY": 230, "Width": 50, "Height": 30, "Dropdown": False, "Step": 0, "MultiSelection": True, "StringItemList": ("First Item", "Second Item", "ThreeItem"), "SelectedItems": (0, 2)})     
    addControlModel("ComboBox", {"PositionX": 160, "PositionY": 230, "Width": 50, "Height": 12, "Dropdown": True, "MaxTextLen": 10, "ReadOnly": False, "StringItemList": ("First Entry", "Second Entry", "Third Entry", "Fourth Entry")} )  
    
    
    dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)  # ダイアログのビューを取得。すべてのコントロールのビューの入れ物。
    dialog.setModel(dialogmodel)  # ビューにモデルを渡す。
    
#     view = dialog.getView()
#     dialogcontroller = Controller(dialog)  # ダイアログのコントローラーを生成。ビューを渡す。



#     combobox1 = dialog.getControl("ComboBox1")
#     print(combobox1.Text)

#     dialog.setVisible(True)
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    dialog.createPeer(toolkit, None)
#     peer = unodialog.getPeer()
#     activateProgressBar(dialog)
    dialog.execute()
    dialog.dispose()
 
 
def controlmodelCreator(dialogmodel):  # Proxyパターンでインスタンスにメソッドを追加する。
    def addControlModel(controltype, props):
        if not "Name" in props:
            props["Name"] = _generateSequentialName(controltype)
        controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))
        values = props.values()
        if any(map(isinstance, values, [tuple]*len(values))):
            [setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
        else:
            controlmodel.setPropertyValues(tuple(props.keys()), tuple(values))
        dialogmodel.insertByName(props["Name"], controlmodel)
    def _generateSequentialName(controltype):
        i = 1
        flg = True
        while flg:
            name = "{}{}".format(controltype, i)
            flg = dialogmodel.hasByName(name)
            i += 1
        return name  
    return addControlModel

 
 
 
 
 
# class ProxyforDialogModel:  # Proxyパターンでインスタンスにメソッドを追加する。
#     def __init__(self, obj):  # メソッドを追加するインスタンスを取得。
#         self._obj = obj
#     def addControlModel(self, controltype, props, *, name=None):
#         name = name or self._generateSequentialName(controltype)
#         if not "Name" in props:
#             props["Name"] = name  
#         controlmodel = self._obj.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))
#         [setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。
#         self._obj.insertByName(name, controlmodel)
#     def _generateSequentialName(self, controltype):
#         i = 1
#         flg = True
#         while flg:
#             name = "{}{}".format(controltype, i)
#             flg = self._obj.hasByName(name)
#             i += 1
#         return name        
#     def __getattr__(self, name):  # Proxyクラス属性にnameが見つからなかったときにnameを引数にして呼び出されます。__setattr__()や __delattr__()が常に呼び出されるのとは対照的です。
#         return getattr(self._obj, name)  # Proxyクラスのインスタンスが取得したインスタンスの属性としてnameを呼び出す。
#     def __setattr__(self, name, value):  # アンダースコアが始まる属性名のときはProxyの属性にvalueを代入し、そうでない時はProxyクラスのインスタンスが取得したインスタンスの属性にvalueを代入する。
#         super().__setattr__(name, value) if name.startswith('_') else setattr(self._obj, name, value)
#     def __delattr__(self, name):  # アンダースコアが始まる属性名のときはProxyの属性を削除し、そうでない時はProxyクラスのインスタンスが取得したインスタンスの属性を削除する。
#         super().__delattr__(name) if name.startswith('_') else delattr(self._obj, name)   

 
 
    
def activateProgressBar(dialog):
    progressbarcontroller = dialog.getControl("ProgressBar1")
    progressbarmodel = progressbarcontroller.getModel()
    progressbarmodel.BackgroundColor = 0x0000ff
    import time
    for i in range(0, 101, 20):
        progressbarmodel.ProgressValue = i
        time.sleep(1)
# class ControlCreator:
#     def __init__(self, dialogmodel, dialogcontrol):
#         dialogcontrol.setModel(dialogmodel)   
#         self.model = dialogmodel
#         self.controller = dialogcontrol
#     def add(self, controltype, props, *, name=None, listeners=None):
#         name = name or self.generateSequentialName(controltype)
#         if not "Name" in props:
#             props["Name"] = name  
#         controlmodel = self.model.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))
#         [setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。
# #         controlmodel.setPropertyValues(tuple(props.keys()), tuple(props.values()))  # タプルが値の時はすべて[]anyと判断されるのでエラーが出る。
#         self.model.insertByName(name, controlmodel)
#         if listeners is not None:
#             controlview = self.controller.getControl(name)
#             for key, val in listeners.items():
#                 getattr(controlview, "add{}".format(key))(val)
#     def generateSequentialName(self, controltype):
#         i = 1
#         flg = True
#         while flg:
#             name = "{}{}".format(controltype, i)
#             flg = self.model.hasByName(name)
#             i += 1
#         return name

class Controller:
    def __init__(self, dialog):
        self.view = dialog
        


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
    def __init__(self, dialogcontrol):
        self.view = dialogcontrol
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