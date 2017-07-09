#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
from com.sun.star.awt import Rectangle
from com.sun.star.beans import NamedValue
import unohelper
from com.sun.star.frame import XController, XTitle, XDispatchProvider
from com.sun.star.lang import XServiceInfo
from com.sun.star.task import XStatusIndicatorFactory

def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
    taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
    args = NamedValue("PosSize", Rectangle(102, 41, 380, 380)), 
    frame = taskcreator.createInstanceWithArguments(args)
    window = frame.getContainerWindow()
    desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    frame.setTitle("Created by TaskCreator")
    frame.setCreator(desktop)
    desktop.getFrames().append(frame)

    
    
#     controlcontainermodel = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlContainerModel', ctx)
#     addControlModel = controlmodelCreator(controlcontainermodel)
#     addControlModel("FixedText", {"Name": "Headerlabel", "PositionX": 106, "PositionY": 6, "Width": 300, "Height": 8, "Label": "This code-sample demonstrates how to create various controls in a dialog"})
#    
#     
#     
#     controlcontainer = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlContainer', ctx)
#     controlcontainer.setModel(controlcontainermodel)
#     controlcontainer.createPeer(toolkit,window)
    
    
    
    window.setVisible(True)


#     frame.activate()

class Controller(unohelper.Base, XController, XTitle, XDispatchProvider, XStatusIndicatorFactory, XServiceInfo):
    IMPLE_NAME = "taskcreater_exam"
    def __init__(self,frame, model):
        self.frame = frame
        self.model = model
    # XTitle
    def getTitle(self):
        return self.frame.getTitle()
    def setTitle(self, title):
        self.frame.setTitle(title)
    def dispose(self):
        self.frame = None
        self.model = None
    def addEventListener(self, xListener):
        pass
    def removeEventListener(self, aListener):
        pass
    # XController
    def attachFrame(self, frame):
        self.frame = frame
    def attachModel(self, model):
        self.model = model
    def suspend(self, Suspend):
        return True
    def getViewData(self):
        """ Returns current instance inspected. """
        return self.ui.main.current.target
    def restoreViewData(self, Data):
        pass
    def getModel(self):
        return self.model
    def getFrame(self):
        return self.frame
    # XStatusIndicatorFactory
    def createStatusIndicator(self):
        pass
    # XDispatchProvider
    def queryDispatch(self, url, name, flags):
        pass
    def queryDispatches(self, requests):
        pass
    # XServiceInfo
    def getImplementationName(self):
        return self.IMPLE_NAME
    def supportsService(self, name):
        return name == self.IMPLE_NAME
    def getSupportedServiceNames(self):
        return self.IMPLE_NAME,




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