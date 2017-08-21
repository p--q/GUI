#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.awt import WindowDescriptor
from com.sun.star.awt import Rectangle
from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.awt.WindowAttribute import  SHOW, BORDER
from com.sun.star.beans import PropertyValue
def macro():
    ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
    smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
    doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
    docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
    docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
    toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。  
    dialog, addControl = dialogCreator(ctx, smgr, {"Name": "Dialog1", "PositionX": 102, "PositionY": 41, "Width": 300, "Height": 400, "Title": "Document-Dialog", "Moveable": True, "TabIndex": 0})  # UnoControlDialogを生成、とそれにコントロールを使いする関数addControl。
    addControl("FixedText", {"Name": "Headerlabel", "PositionX": 6, "PositionY": 6, "Width": 300, "Height": 8, "Label": "This code-sample demonstrates how to display an office document in a dialog window", "NoLabel": True})
    addControl("Button", {"PositionX": 126, "PositionY": 370, "Width": 50, "Height": 14, "Label": "~Close dialog", "PushButtonType": 1})  # PushButtonTypeの値はEnumではエラーになる。
    dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。
    dialogwindow = dialog.getPeer()  # ダイアログウィンドウ(=ピア）を取得。
    subwindow =  createWindow(toolkit, SHOW + BORDER, {"PositionX": 40, "PositionY": 50, "Width": 420, "Height": 550, "ParentIndex": 1, "Parent": dialogwindow, "WindowServiceName": "dockingwindow", "Type": SIMPLE})  # ツールキットを使ってドキュメントウィンドウの上にウィンドウを作成する。3番目の引数サービス名はcom.sun.star.awt.WindowDescriptorで定義されている。
    subframe = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しいフレームを生成。
    subframe.initialize(subwindow)  # フレームにコンテナウィンドウを入れる。  
    nodes = PropertyValue(Name = "Preview", Value = True), PropertyValue(Name = "ReadOnly", Value = True)  # com.sun.star.document.MediaDescriptor
    subframe.loadComponentFromURL("private:factory/swriter", "_self", 2, nodes) # フレームのコンポーネントウィンドウにWriterドキュメントをロード。
    # ノンモダルダイアログにするとき。
#     showModelessly(ctx, smgr, docframe, dialog)  
    # モダルダイアログにする。フレームに追加するとエラーになる。
    dialog.execute()  
    dialog.dispose()   
def createWindow(toolkit, attr, props):  # ウィンドウタイトルは変更できない。attrはcom.sun.star.awt.WindowAttributeの和。propsはPositionX, PositionY, Width, Height, ParentIndex, Parent, WindowServiceName, Type。
    aRect = Rectangle(X=props.pop("PositionX"), Y=props.pop("PositionY"), Width=props.pop("Width"), Height=props.pop("Height"))
    d = WindowDescriptor(Bounds=aRect, WindowAttributes=attr)
    for key, val in props.items():
        setattr(d, key, val)
    return toolkit.createWindow(d)  # ウィンドウピアを返す。 
def showModelessly(ctx, smgr, parentframe, dialog):  # ノンモダルダイアログにする。オートメーションではリスナー動かない。ノンモダルダイアログではフレームに追加しないと閉じるボタンが使えない。
    frame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", ctx)  # 新しいフレームを生成。
    frame.initialize(dialog.getPeer())  # フレームにコンテナウィンドウを入れる。    
    frame.setName(dialog.getModel().getPropertyValue("Name"))  # フレーム名をダイアログモデル名から取得（一致させる必要性はない）して設定。
    parentframe.getFrames().append(frame)  # 新しく作ったフレームを既存のフレームの階層に追加する。 
    dialog.setVisible(True)  # ダイアログを見えるようにする。   
    return frame  # フレームにリスナーをつけるときのためにフレームを返す。
def dialogCreator(ctx, smgr, dialogprops):  # ダイアログと、それにコントロールを追加する関数を返す。まずダイアログモデルのプロパティを取得。
    dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)  # ダイアログの生成。
    if "PosSize" in dialogprops:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
        dialog.setPosSize(dialogprops.pop("PositionX"), dialogprops.pop("PositionY"), dialogprops.pop("Width"), dialogprops.pop("Height"), dialogprops.pop("PosSize"))  # ダイアログモデルのプロパティで設定すると単位がMapAppになってしまうのでコントロールに設定。
    dialogmodel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)  # ダイアログモデルの生成。
    dialogmodel.setPropertyValues(tuple(dialogprops.keys()), tuple(dialogprops.values()))  # ダイアログモデルのプロパティを設定。
    dialog.setModel(dialogmodel)  # ダイアログにダイアログモデルを設定。
    dialog.setVisible(False)  # 描画中のものを表示しない。
    def addControl(controltype, props, attrs=None):  # props: コントロールモデルのプロパティ、attr: コントロールの属性。
        items, currentitemid = None, None
        if controltype == "Roadmap":  # Roadmapコントロールのとき、Itemsはダイアログモデルに追加してから設定する。そのときはCurrentItemIDもあとで設定する。
            if "Items" in props:  # Itemsはダイアログモデルに追加されてから設定する。
                items = props.pop("Items")
                if "CurrentItemID" in props:  # CurrentItemIDはItemsを追加されてから設定する。
                    currentitemid = props.pop("CurrentItemID")
        if "PosSize" in props:  # コントロールモデルのプロパティの辞書にPosSizeキーがあるときはピクセル単位でコントロールに設定をする。
            control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl{}".format(controltype), ctx)  # コントロールを生成。
            control.setPosSize(props.pop("PositionX"), props.pop("PositionY"), props.pop("Width"), props.pop("Height"), props.pop("PosSize"))  # ピクセルで指定するために位置座標と大きさだけコントロールで設定。
            controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
            control.setModel(controlmodel)  # コントロールにコントロールモデルを設定。
            dialog.addControl(props["Name"], control)  # コントロールをコントロールコンテナに追加。
        else:  # Map AppFont (ma)のときはダイアログモデルにモデルを追加しないと正しくピクセルに変換されない。
            controlmodel = _createControlModel(controltype, props)  # コントロールモデルの生成。
            dialogmodel.insertByName(props["Name"], controlmodel)  # ダイアログモデルにモデルを追加するだけでコントロールも作成される。
        if items is not None:  # コントロールに追加されたRoadmapモデルにRoadmapアイテムは追加できない。
            for i, j in enumerate(items):  # 各Roadmapアイテムについて
                item = controlmodel.createInstance()
                item.setPropertyValues(("Label", "Enabled"), j)
                controlmodel.insertByIndex(i, item)  # IDは0から整数が自動追加される       
            if currentitemid is not None:  #Roadmapアイテムを追加するとそれがCurrentItemIDになるので、Roadmapアイテムを追加してからCurrentIDを設定する。
                controlmodel.setPropertyValue("CurrentItemID", currentitemid)
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
    def _generateSequentialName(controltype):  # コントロールの連番名の作成。
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
    from functools import wraps
    import sys
#     from com.sun.star.beans import PropertyValue
    from com.sun.star.script.provider import XScriptContext  
    def connectOffice(func):  # funcの前後でOffice接続の処理
        @wraps(func)
        def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
            try:
                ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
            except:
                print("Could not establish a connection with a running office.", file=sys.stderr)
                sys.exit()
            print("Connected to a running office ...")
            smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
            print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
            return func(ctx, smgr)  # 引数の関数の実行。
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
                return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
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