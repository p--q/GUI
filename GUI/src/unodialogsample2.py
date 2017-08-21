#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.beans.MethodConcept import ALL as MethodConcept_ALL
from com.sun.star.beans.PropertyConcept import ATTRIBUTES as PropertyConcept_ATTRIBUTES, PROPERTYSET as PropertyConcept_PROPERTYSET
from com.sun.star.awt import XItemListener
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
    def wrapper(*args, **kwargs):
        import pydevd
        frame = None
        doc = XSCRIPTCONTEXT.getDocument()
        if doc:  # ドキュメントが取得できた時
            frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。
        else:
            currentframe = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()  # モードレスダイアログのときはドキュメントが取得できないので、モードレスダイアログのフレームからCreatorのフレームを取得する。
            frame = currentframe.getCreator()
        if frame:   
            import time
            indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
            maxrange = 2  # ステータスバーに表示するプログレスバーの目盛りの最大値。2秒ロスするが他に適当な告知手段が思いつかない。
            indicator.start("Trying to connect to the PyDev Debug Server for about 20 seconds.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
            t = 1  # プレグレスバーの初期値。
            while t<=maxrange:  # プログレスバーの最大値以下の間。
                indicator.setValue(t)  # プレグレスバーの位置を設定。
                time.sleep(1)  # 1秒待つ。
                t += 1  # プログレスバーの目盛りを増やす。
            indicator.end()  # reset()の前にend()しておかないと元に戻らない。
            indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。
        pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。
        try:
            func(*args, **kwargs)  # Step Intoして中に入る。
        except:
            import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
    return wrapper
# @enableRemoteDebugging
def macro():
    ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
    smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
    doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
    docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
    docwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
    toolkit = docwindow.getToolkit()  # ピアからツールキットを取得。  
    controlmargin = 6  # ロードマップコントロール以外のコントロールのマージン。
    dialogwidth, dialogheight = 250, 140  # ダイアログの幅と高さ
    roadmapwidth = 80  # ロードマップの幅。
    buttonwidth, buttonheight = 50, 14  # ボタンコントロールの幅と高さ。
    buttonposx = int(dialogwidth/2 - buttonwidth/2)  # 整数に変換が必要。ボタンの位置をダイアログの中央にもってくる。
    buttonposy = dialogheight - buttonheight - controlmargin  # ボタンのY座標。下縁からcontrolmarginの位置にする。
    controlposx = roadmapwidth + controlmargin  # Stepで切り替えるコントロールのX座標。ロードマップの右縁からcontrolmarginを確保。
    controlwidth = dialogwidth - 2*controlmargin - roadmapwidth  # Stepで切り替えるコントロールの幅。左右にcontrolmarginを確保。
    listboxheight = dialogheight - 4*controlmargin - buttonheight  # リストボックスの高さ。ボタンのcontrolmarginも引く。
    dialog, addControl = dialogCreator(ctx, smgr, {"Name": "Dialog1", "PositionX": 102, "PositionY": 41, "Width": dialogwidth, "Height": dialogheight, "Title": "Inspect a Uno-Object", "Moveable": True, "TabIndex": 0, "Step": 1})  # UnoControlDialogを生成、とそれにコントロールを使いする関数addControl。最初に表示するStepを指定している。
    linecount, fixedtextheight = 4, 8  # FixedTextの行数、1行の高さ。
    label = "This Dialog lists information about a given Uno-Object.\nIt offers a view to inspect all suppported servicenames, exported interfaces, methods and properties."
    addControl("FixedText", {"PositionX": controlposx, "PositionY": 27, "Width": controlwidth, "Height": fixedtextheight*linecount, "Label": label, "NoLabel": True, "Step": 1, "MultiLine": True})  # Step1で表示させるコントロール。
    introspection = smgr.createInstanceWithContext("com.sun.star.beans.Introspection", ctx)  # ロードマップコントロールで切り替えるネタの取得のため。
    introspectionaccess = introspection.inspect(doc)  # ドキュメントモデルを調べる。
    supportedservicenames = doc.getSupportedServiceNames()  # ドキュメントモデルのサービス一覧を取得。
    interfacenames = tuple(i.typeName for i in doc.getTypes())  # ドキュメントモデルのインターフェイス一覧を取得。 getTypename()メソッドはないと言われる。
    methodnames = tuple(i.getName() for i in introspectionaccess.getMethods(MethodConcept_ALL))  # ドキュメントモデルのメソッド一覧を取得。
    propertynames = tuple(i.Name for i in introspectionaccess.getProperties(PropertyConcept_ATTRIBUTES + PropertyConcept_PROPERTYSET))  # ドキュメントモデルのプロパティ一覧を取得。
    addControl("ListBox", {"PositionX": controlposx, "PositionY": controlmargin, "Width": controlwidth, "Height": listboxheight, "Dropdown": False, "ReadOnly": True, "Step": 2, "StringItemList": supportedservicenames})  # Step2で表示させるコントロール。     
    addControl("ListBox", {"PositionX": controlposx, "PositionY": controlmargin, "Width": controlwidth, "Height": listboxheight, "Dropdown": False, "ReadOnly": True, "Step": 3, "StringItemList": interfacenames})  # Step3で表示させるコントロール。     
    addControl("ListBox", {"PositionX": controlposx, "PositionY": controlmargin, "Width": controlwidth, "Height": listboxheight, "Dropdown": False, "ReadOnly": True, "Step": 4, "StringItemList": methodnames})  # Step4で表示させるコントロール。     
    addControl("ListBox", {"PositionX": controlposx, "PositionY": controlmargin, "Width": controlwidth, "Height": listboxheight, "Dropdown": False, "ReadOnly": True, "Step": 5, "StringItemList": propertynames})  # Step5で表示させるコントロール。     
    addControl("Button", {"PositionX": buttonposx, "PositionY": buttonposy, "Width": buttonwidth, "Height": buttonheight, "Label": "~Close", "PushButtonType": 1})  # ダイアログを閉じるボタンのコントロール。PushButtonTypeの値はEnumではエラーになる。
    addControl("FixedLine", {"PositionX": 0, "PositionY": buttonposy - controlmargin - 4, "Width": dialogwidth, "Height": 8, "Orientation": 0})  # 水平線の高さをロードマップの下縁の半分に食い込ませる。
    dialog.createPeer(toolkit, docwindow)  # ダイアログを描画。親ウィンドウを渡す。ノンモダルダイアログのときはNone(デスクトップ)ではフリーズする。Stepを使うときはRoadmap以外のコントロールが追加された後にピアを作成しないとStepが重なって表示される。
    items = ("Introduction", True),\
            ("Supported Services", True),\
            ("Interfaces", True),\
            ("Methods", True),\
            ("Properties", True)  # この順に0からIDがふられる。この順に表示される。
    addControl("Roadmap", {"PositionX": 0, "PositionY": 0, "Width": roadmapwidth, "Height": dialogheight - buttonheight - 2*controlmargin, "Complete": True, "CurrentItemID": 0, "Text": "Steps", "Items": items}, {"addItemListener": ItemListener(dialog)})  # Roadmapコントロールはダイアログウィンドウを描画してからでないと項目が表示されない。
    # ノンモダルダイアログにするとき。オートメーションではリスナーが動かない。
#     showModelessly(ctx, smgr, docframe, dialog)  
    # モダルダイアログにする。フレームに追加するとエラーになる。
    dialog.execute()  
    dialog.dispose()   
class ItemListener(unohelper.Base, XItemListener): 
    def __init__(self, dialog):
        self.dialogmodel = dialog.getModel()
#     @enableRemoteDebugging
    def itemStateChanged(self, itemevent):
        dummy_control, dummy_controlmodel, name = eventSource(itemevent)
        if name == "Roadmap1":  # ロードマップコントロールのアイテムの選択が変更されたとき
            itemid = itemevent.ItemId + 1  # ItemIdは0から始まるのでStepと合わせるため1足す。
            step = self.dialogmodel.getPropertyValue("Step")  # 現在のダイアログのStepを取得。
            if itemid != step:  # ItemIdが現在のStepと異なるとき。
                self.dialogmodel.setPropertyValue("Step", itemid)  # StepをItemIdにする。
    def disposing(self, eventobject):
        pass 
def eventSource(event):  # イベントからコントロール、コントロールモデル、コントロール名を取得。
    control = event.Source  # イベントを駆動したコントロールを取得。
    controlmodel = control.getModel()  # コントロールモデルを取得。
    name = controlmodel.getPropertyValue("Name")  # コントロール名を取得。    
    return control, controlmodel, name    
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
    from com.sun.star.beans import PropertyValue
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
