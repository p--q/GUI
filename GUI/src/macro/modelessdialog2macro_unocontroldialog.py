#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import uno  # オートメーションのときのみ必要。
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
    dialog, addControl = dialogCreator(ctx, smgr, {"PositionX": 150, "PositionY": 150, "Width": 200, "Height": 200, "Title": "New Dialog", "Name": "dialog", "Step": 0, "TabIndex": 0, "Moveable": True})
    dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。
    dialogwindow = dialog.getPeer()  # ダイアログウィンドウを取得。
    addControl("FixedText", {"PositionX": 10, "PositionY": 0, "Width": 180, "Height": 30, "Label": "~Label1", "VerticalAlign": BOTTOM})
    addControl("Edit", {"PositionX": 10, "PositionY": 40, "Width": 180, "Height": 30}, {"setFocus": None})
    addControl("Button", {"PositionX": 110, "PositionY": 130, "Width": 80, "Height": 35, "DefaultButton": True, "Label": "~btn"}, {"setActionCommand": "btn", "addActionListener": BtnListener(dialog, dialogwindow)})
    # ノンモダルダイアログ。オートメーションではリスナー動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。
    createFrame = frameCreator(ctx, smgr, docframe)  # 親フレームを渡す。
    createFrame("NewFrame", dialogwindow)  # 新しいフレーム名、そのコンテナウィンドウ。
    dialog.setVisible(True)  # ダイアログを見えるようにする。
    # ここで止まるとノンモダルダイアログになる。オートメーションではリスナー動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。
    dialog.execute()  # モダルダイアログにする。
    dialog.dispose()    
class BtnListener(unohelper.Base, XActionListener):  # ボタンリスナー
    def __init__(self, dialog, dialogwindow):  # dialogwindowはメッセージボックス表示のため。  
        self.dialog = dialog  # ダイアログを取得。
        self.window = dialogwindow  # ダイアログウィンドウを取得。
    def actionPerformed(self, actionevent):
        cmd = actionevent.ActionCommand  # アクションコマンドを取得。
        edit = self.dialog.getControl("Edit1")  # editという名前のコントロールを取得。
        if cmd == "btn":  # アクションコマンドがbtnのとき
            edit.setText("By Button Click")  # editコントロールに文字列を代入。
            toolkit = self.window.getToolkit()  # ツールキットを取得。
            msgbox = toolkit.createMessageBox(self.window, INFOBOX, BUTTONS_OK, "Text Field", "{}".format(edit.getText()))  # self.windowを親ウィンドウにしてメッセージボックスを作成。
            msgbox.execute()  # メッセージボックスを表示。
            msgbox.dispose()  # メッセージボックスを破棄。
    def disposing(self, eventobject):
        pass  
def frameCreator(ctx, smgr, parentframe): # 新しいフレームを追加する関数を返す。親フレームを渡す。   
    def createFrame(framename, containerwindow):  # 新しいフレーム名、そのコンテナウィンドウにするウィンドウを渡す。
        frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しく作成したウィンドウを入れるためのフレームを作成。
        frame.initialize(containerwindow)  # フレームにウィンドウを入れる。    
        frame.setName(framename)  # フレーム名を設定。
        parentframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。 
        return frame        
    return createFrame               
def dialogCreator(ctx, smgr, dialogprops):  # ダイアログと、それにコントロールを追加する関数を返す。まずダイアログモデルのプロパティを取得。
    dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)  # ダイアログの生成。
    dialog.setPosSize(dialogprops.pop("PositionX"), dialogprops.pop("PositionY"), dialogprops.pop("Width"), dialogprops.pop("Height"), POSSIZE)  # ダイアログモデルのプロパティで設定すると単位がMapAppになってしまうのでコントロールに設定。
    dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログモデルの生成。
    dialogmodel.setPropertyValues(tuple(dialogprops.keys()), tuple(dialogprops.values()))  # ダイアログモデルのプロパティを設定。
    dialog.setModel(dialogmodel)  # ダイアログにダイアログモデルを設定。
    dialog.setVisible(False)  # 描画中のものを表示しない。
    def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
        control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
        control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), POSSIZE)  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
        if not "Name" in props:
            props["Name"] = _generateSequentialName(controltype)  # Nameがpropsになければ通し番号名を生成。
        controlmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}Model".format(controltype), ctx)  # コントロールモデルを生成。
        values = props.values()  # プロパティの値がタプルの時にsetProperties()でエラーが出るのでその対応が必要。
        if any(map(isinstance, values, [tuple]*len(values))):
            [setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
        else:
            controlmodel.setPropertyValues(tuple(props.keys()), tuple(values))
        control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
        dialog.addControl(props["Name"], control)  # コントロールをダイアログに追加。
        if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
            control = dialog.getControl(props["Name"])  # ダイアログに追加された後のコントロールを取得。
            for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
                if val is None:
                    getattr(control, key)()
                else:
                    getattr(control, key)(val)
    def _generateSequentialName(controltype):  # 連番名の作成。
        i = 1
        flg = True
        while flg:
            name = "{}{}".format(controltype, i)
            flg = dialog.getControl(name)  # 同名のコントロールの有無を判断。
            i += 1
        return name  
    return dialog, addControl  # ダイアログとそのダイアログにコントロールを追加する関数を返す。
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