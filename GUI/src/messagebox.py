#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
from com.sun.star.awt.MessageBoxType import ERRORBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.util import CloseVetoException
from com.sun.star.uno import Exception
from com.sun.star.lang import IllegalArgumentException
def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
    model = None
    try:
        messagebox = MessageBox(ctx, smgr)
        model = messagebox.createDefaultTextDocument()
        peer = messagebox.getWindowPeerOfFrame(model)
        if peer is not None:
            bishighcontrast = messagebox.isHighContrastModeActivated(peer)
            messagebox.showErrorMessageBox(peer, "My Sampletitle", "HighContrastMode is enabled: {}".format(bishighcontrast))
        else:
            print("Could not retrieve current frame")
    except Exception as e:
        print(e + e.getMessage())
        traceback.print_exc()
    finally:
        if model is not None:
            try:
                if hasattr(model, "close"):
                    model.close(False)
                else:
                    model.dispose()
            except CloseVetoException as e:
                print(e + e.getMessage())
                traceback.print_exc()
class MessageBox:
    def __init__(self, ctx, smgr):
        self.ctx = ctx
        self.smgr = smgr
    def getWindowPeerOfFrame(self, model):
        try:
            frame = None
            if model is not None:
                frame = model.getCurrentController().getFrame()
            else:
                desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
                frame = desktop.getActiveFrame()
            if frame is not None:
                return frame.getContainerWindow()
        except Exception:  
            traceback.print_exc()
        return None
    def createDefaultTextDocument(self):
        model = None
        try:
            desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
            prop = PropertyValue(Name="Hidden", Value=True)
            model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))
        except Exception:
            traceback.print_exc()
        return model
    def showErrorMessageBox(self, parentwindowpeer, title, message):
        try:
            toolkit = self.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
            messagebox = toolkit.createMessageBox(parentwindowpeer, ERRORBOX, BUTTONS_OK, title, message)
            if messagebox is not None:
                messagebox.execute()
        except Exception:
            traceback.print_exc()
        finally:
            if messagebox is not None:
                messagebox.dispose()
    def isHighContrastModeActivated(self, vclwindowpeer):
        bisactivated = False
        try:
            if vclwindowpeer is not None:
                color = vclwindowpeer.getProperty("DisplayBackgroundColor")
                red = self.getRedColorShare(color)
                green = self.getGreenColorShare(color)
                blue = self.getBlueColorShare(color)
                luminance = (blue*28 + green*151 + red*77 ) / 256
                bisactivated = (luminance <= 25)
                return bisactivated
            else:
                return False
        except IllegalArgumentException:
            traceback.print_exc()
        return bisactivated
    def getRedColorShare(self, color):
        return color/65536
    def getGreenColorShare(self, color):
        redmodulo = color%65536
        return redmodulo/256
    def getBlueColorShare(self, color):
        redmodulo = color%65536
        return redmodulo%256
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
    