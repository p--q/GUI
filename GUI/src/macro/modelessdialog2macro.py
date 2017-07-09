#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
# from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.style.VerticalAlignment import BOTTOM
import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.awt import Rectangle
from com.sun.star.awt.MessageBoxType import INFOBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.beans import NamedValue

def macro():
    ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
    smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
    doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
    docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
    docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウを取得。
    toolkit = docwindow.getToolkit()  # ツールキットを取得。
    taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
    args = NamedValue("PosSize", Rectangle(150, 150, 200, 200)), NamedValue("FrameName", "NewFrame"), NamedValue("MakeVisible", True)
    frame = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
    subwindow = frame.getContainerWindow()  # 新しいコンテナウィンドウを新しいフレームから取得。
    frame.setTitle("New Frame")  # フレームのタイトルを設定。
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
#     subwindow.setVisible(True)  # 新しく作ったウィンドウを見えるようにする。これがなくても表示される。     
#     subwindow.execute()  # TaskCreatorから得たコンテナウィンドウにはexecute()がない。
#     subwindow.dispose()
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
def createControl(smgr, ctx, ctype, x, y, width, height, names, values):
    ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(ctype), ctx)
    ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}Model".format(ctype), ctx)
    ctrl_model.setPropertyValues(names, values)
    ctrl.setModel(ctrl_model)
    ctrl.setPosSize(x, y, width, height, POSSIZE)
    return ctrl
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。