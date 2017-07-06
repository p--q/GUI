#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue

def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
    from msgbox import MsgBox
    myBox = MsgBox(ctx)
    myBox.addButton("OK")
    myBox.renderFromBoxSize(100)  # メッセージボックスの幅を文字列の長さから算出。
    myBox.show("message", 0, "Error")


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