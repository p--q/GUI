#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.style.VerticalAlignment import BOTTOM
import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.awt.MessageBoxType import INFOBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.awt.WindowAttribute import  CLOSEABLE, SHOW, MOVEABLE, BORDER
from com.sun.star.awt import WindowDescriptor
from com.sun.star.awt import Rectangle
from com.sun.star.beans import NamedValue


def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
    docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
    docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウを取得。
    toolkit = docwindow.getToolkit()  # ツールキットを取得。
    subwindow =  createWindow(toolkit, docwindow, "dialog", SHOW + BORDER + MOVEABLE + CLOSEABLE, 150, 150, 200, 200)  # ツールキットを使ってドキュメントウィンドウの上にウィンドウを作成する
    frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しく作成したウィンドウを入れるためのフレームを作成。
    frame.initialize(subwindow)  # フレームにウィンドウを入れる。
#     frame.setCreator(docframe)  # フレームの親フレームを設定する。
    frame.setName("NewFrame")  # フレーム名を設定。
#     frame.setTitle("New Frame")  # フレームのタイトルを設定。これはバグで反映されない。
    docframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。
    controlcontainer = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)  # コントロールの集合を作成。
    controlcontainermodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)  # コントールのモデルの集合を作成。
    controlcontainermodel.setPropertyValue("BackgroundColor", -1)  # 背景色。-1は何色?
    controlcontainer.setModel(controlcontainermodel)  # コントールコンテナにモデルを設定。
    controlcontainer.createPeer(toolkit, subwindow)  # 新しく作ったウィンドウ内にコントロールコンテナのコントロールをツールキットで描画する。
    controlcontainer.setPosSize(0, 0, 200, 200, POSSIZE)  # コントロールの表示座標を設定。4つ目の引数は前の引数の意味を設定する。
    frame.setComponent(controlcontainer, None)  # フレームにコントローラを設定する。今回のコントローラはNone。
    label = createControl(smgr, ctx, "FixedText", 10, 0, 180, 30, ("Label", "VerticalAlign"), ("Label1", BOTTOM))  # 固定文字コントールを作成。
    edit = createControl(smgr, ctx, "Edit", 10, 40, 180, 30, (), ())  # 編集枠コントロールを作成。
    btn = createControl(smgr, ctx, "Button", 110, 130, 80, 35, ("DefaultButton", "Label"), (True, "btn"))  # ボタンコントロールを作成。
    controlcontainer.addControl("label", label)  # コントロールコンテナにコントロールを名前を設定して追加。
    controlcontainer.addControl("btn", btn)  # コントロールコンテナにコントロールを名前を設定して追加。
    controlcontainer.addControl("edit", edit)  # コントロールコンテナにコントロールを名前を設定して追加。
    edit.setFocus()  # 編集枠コントロールにフォーカスを設定する。
    btn.setActionCommand("btn")  # ボタンを起動した時のコマンド名を設定する。
    btn.addActionListener(BtnListener(controlcontainer, subwindow))  # ボタンにリスナーを設定。コントロールの集合を渡しておく。
#     subwindow.setVisible(True)  # モードレスダイアログにするとボタンのリスナーが呼ばれない。
    subwindow.execute()  # execute()にするとモダルダイアログになる。
    subwindow.dispose()
class BtnListener(unohelper.Base, XActionListener):
    def __init__(self, controlcontainer, window):  
        self.controlcontainer = controlcontainer  # コントロールの集合。
        self.window = window  # コントロールのあるウィンドウ。
    def actionPerformed(self, actionevent):
        cmd = actionevent.ActionCommand  # アクションコマンドを取得。
        editcontrol = self.controlcontainer.getControl("edit")  # editという名前のコントロールを取得。
        if cmd == "btn":  # 開くsyんコマンドがbtnのとき
            editcontrol.setText("By Button Click")  # editコントロールに文字列を代入。
            toolkit = self.window.getToolkit()
            msgbox = toolkit.createMessageBox(self.window, INFOBOX, BUTTONS_OK, "Text Field", "{}".format(editcontrol.getText()))  # ピアオブジェクトからツールキットを取得して、peerを親ウィンドウにしてメッセージボックスを作成。
            msgbox.execute()  # メッセージボックスを表示。
            msgbox.dispose()  # メッセージボックスを破棄。
    def disposing(self, eventobject):
        pass              
def createWindow(toolkit, parent, service, attr, nX, nY, nWidth, nHeight):
    aRect = Rectangle(X=nX, Y=nY, Width=nWidth, Height=nHeight)
    d = WindowDescriptor(Type=SIMPLE, WindowServiceName=service, ParentIndex=-1, Bounds=aRect, Parent=parent, WindowAttributes=attr)
    return toolkit.createWindow(d)
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