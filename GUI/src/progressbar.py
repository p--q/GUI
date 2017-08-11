#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
def macro():
    ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
    smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
    doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
    docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
    docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウを取得。
    toolkit = docwindow.getToolkit()  # ツールキットを取得。  
    dialog, addControl = dialogCreator(ctx, smgr, {"PositionX": 102, "PositionY": 41, "Width": 380, "Height": 380, "Title": "LibreOffice", "Name": "MyTestDialog", "Step": 0, "TabIndex": 0, "Moveable": True, "Tag": "ダイアログ"})
    dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。ツールキットと親ウィンドウを渡す。親ウィンドウがNoneのときはdesktopになる。
    addControl("ProgressBar", {"PositionX": 106, "PositionY": 44, "Width": 250, "Height": 8, "ProgressValueMin": 0, "ProgressValueMax": 100})
    # ノンモダルダイアログ。オートメーションではリスナー動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。リスナー以外もオートメーションでは動かない時あり。
    dialogwindow = dialog.getPeer()  # ダイアログウィンドウを取得。
    createFrame = frameCreator(ctx, smgr, docframe)  # 親フレームを渡す。
    createFrame("NewFrame", dialogwindow)  # 新しいフレーム名、そのコンテナウィンドウ。


    dialog.setVisible(True)  # ダイアログを見えるようにする。
    activateProgressBar(toolkit, dialog, docwindow)

    # モダルダイアログにする。フレームに追加するとエラーになる。
#     dialog.execute()  
#     dialog.dispose()    


#             doc = XSCRIPTCONTEXT.getDocument()
#             if doc:
#                 docframe = doc.getCurrentController().getFrame() 
#                 indicator = docframe.createStatusIndicator()
#                 indicatormax = 50
#                 indicator.start("Waiting for establishing a connection with PyDev Debug Server :  ", indicatormax)
#                 t = 0
#                 while t<indicatormax:
#                     indicator.setValue(t)
#                     time.sleep(1)
#                     t += 1
#                 indicatormax.end()



def activateProgressBar(toolkit, dialog, docwindow):
    progressbarcontrol = dialog.getControl("ProgressBar1")
    progressbarmodel = progressbarcontrol.getModel()
    progressbarmodel.BackgroundColor = 0x0000ff
    import time
    for i in range(0, 101, 20):
        progressbarmodel.ProgressValue = i
        dialog.createPeer(toolkit, docwindow)
#         progressbarcontrol.createPeer(toolkit, dialogwindow)
        time.sleep(1)
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
    if "PosSize" in dialogprops:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
        dialog.setPosSize(dialogprops.pop("PositionX"), dialogprops.pop("PositionY"), dialogprops.pop("Width"), dialogprops.pop("Height"), dialogprops.pop("PosSize"))  # ダイアログモデルのプロパティで設定すると単位がMapAppになってしまうのでコントロールに設定。
    dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログモデルの生成。
    dialogmodel.setPropertyValues(tuple(dialogprops.keys()), tuple(dialogprops.values()))  # ダイアログモデルのプロパティを設定。
    dialog.setModel(dialogmodel)  # ダイアログにダイアログモデルを設定。
    dialog.setVisible(False)  # 描画中のものを表示しない。
    def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
        if "PosSize" in props:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
            control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
            control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
            controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
            control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
            dialog.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
        else:  # Map AppFont (ma)のときはダイアログモデルにモデルを追加しないと正しくピクセルに変換されない。
            controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
            dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
        if attrs is not None:  # Dialogに追加したあとでないと各コントロールへの属性は追加できない。
            control = dialog.getControl(props["Name"])  # コントロールコンテナに追加された後のコントロールを取得。
            for key, val in attrs.items():  # メソッドの引数がないときはvalをNoneにしている。
                if val is None:
                    getattr(control, key)()
                else:
                    getattr(control, key)(val)
    def _createControlModel(controltype, props):  # コントロールモデルの生成。
        if not "Name" in props:
            props["Name"] = _generateSequentialName(controltype)  # Nameがpropsになければ通し番号名を生成。
        controlmodel = dialogmodel.createInstance("com.sun.star.awt.UnoControl{}Model".format(controltype))  # コントロールモデルを生成。UnoControlDialogElementサービスのためにUnoControlDialogModelからの作成が必要。
        if props:
            values = props.values()  # プロパティの値がタプルの時にsetProperties()でエラーが出るのでその対応が必要。
            if any(map(isinstance, values, [tuple]*len(values))):
                [setattr(controlmodel, key, val) for key, val in props.items()]  # valはリストでもタプルでも対応可能。XMultiPropertySetのsetPropertyValues()では[]anyと判断されてタプルも使えない。
            else:
                controlmodel.setPropertyValues(tuple(props.keys()), tuple(values))
        return controlmodel
    def _generateSequentialName(controltype):  # 連番名の作成。
        i = 1
        flg = True
        while flg:
            name = "{}{}".format(controltype, i)
            flg = dialog.getControl(name)  # 同名のコントロールの有無を判断。
            i += 1
        return name
    return dialog, addControl  # コントロールコンテナとそのコントロールコンテナにコントロールを追加する関数を返す。
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