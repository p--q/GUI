#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
from com.sun.star.awt.MessageBoxType import ERRORBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from contextlib import contextmanager

def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    with loadmodel(desktop) as model:
        peer = model.getCurrentController().getFrame().getContainerWindow()
        bishighcontrast = False
        msgbox = peer.getToolkit().createMessageBox(peer, ERRORBOX, BUTTONS_OK, "My Sampletitle", "HighContrastMode is enabled: {}".format(bishighcontrast))
        msgbox.execute()
        msgbox.dispose()
@contextmanager
def loadmodel(desktop):
    prop = PropertyValue(Name="Hidden", Value=True)
    model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,)) 
    yield model
    if hasattr(model, "close"):
        model.close(False)
    else:
        model.dispose()    
    
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
        if not smgr:
            print( "ERROR: no service manager" )
            sys.exit()
        print("Using remote servicemanager\n") 
        try:
            func(ctx, smgr)  # 引数の関数の実行。
        except:
            traceback.print_exc()
        # soffice.binの終了処理。これをしないとLibreOfficeを起動できなくなる。
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        prop = PropertyValue(Name="Hidden", Value=True)
        desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))  # バックグラウンドでWriterのドキュメントを開く。
        terminated = desktop.terminate()  # LibreOfficeをデスクトップに展開していない時はエラーになる。
        if terminated:
            print("\nThe Office has been terminated.")  # 未保存のドキュメントがないとき。
        else:
            print("\nThe Office is still running. Someone else prevents termination.")  # 未保存のドキュメントがあってキャンセルボタンが押された時。
    return wrapper
if __name__ == "__main__":
    main = connectOffice(main)
    main()
    