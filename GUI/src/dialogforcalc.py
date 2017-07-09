# -*- coding: utf-8 -*-
 
# v0.1
import uno
import unohelper
 
from com.sun.star.awt import XActionListener, Rectangle
from com.sun.star.awt.PosSize import POS as PS_POS
from com.sun.star.lang import XComponent
 
 
class DialogBase(unohelper.Base, object):
    """ Helper class to make dialog class. """
    
    def __init__(self, ctx, frame, dialog_url="", string_resource={}):
        self.ctx = ctx
        self._frame = frame
        self.dialog_url = dialog_url
        self.strres = string_resource
        
        self._disposed = False
        self._dialog = None
    
    def _s(self, text):
        """ Translates text. """
        return self.strres.get(text.replace(" ", "_"), text)
    
    # XEventListener
    def disposing(self, ev):
        self.ctx = None
        self._dialog = None
        self._strres = None
        self._frame = None
    
    def _dispose(self):
        """ Destroy the dialog. """
        if self._disposed: return
        self._dialog.dispose()
        self._disposed = True
    
    def centering_dialog(self):
        """ Put the dialog on center of the parent window. """
        if not self._dialog: return
        ps = self._frame.getContainerWindow().getPosSize()
        dps = self._dialog.getPosSize()
        self._dialog.setPosSize(
            (ps.Width - dps.Width)/2, (ps.Height - dps.Height)/2, 0, 0, PS_POS)
        
        x, y = self._dialog.getModel().getPropertyValues(("PositionX", "PositionY"))
        self._dialog.getModel().setPropertyValues(("PositionX", "PositionY"), (x, y))
    
    def _create_dialog(self, x=0, y=0, width=0, height=0, title=""):
        """ Create dialog instance. 
        if dialog_url has passed, the dialog is created from it. 
        Otherwise, runtime dialog creation is started.
        """
        ctx = self.ctx
        smgr = self.ctx.getServiceManager()
        
        if self.dialog_url:
            dp = smgr.createInstanceWithContext(
                "com.sun.star.awt.DialogProvider", ctx)
            dialog = dp.createDialog(self.dialog_url)
        else:
            dialog = smgr.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialog", ctx)
            dialog_model = smgr.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialogModel", ctx)
            dialog_model.setPropertyValues(
                ("Height", "PositionX", "PositionY", "Title", "Width"), 
                (height, x, y, title, width))
            dialog.setModel(dialog_model)
        
        #dialog.setVisible(True)
        #dialog.createPeer(self._parent.getToolkit(), self._parent)
        #dialog.setVisible(False)
        self._dialog = dialog
    
    def _create_peer(self):
        parent = self._frame.getContainerWindow()
        self._dialog.createPeer(parent.getToolkit(), parent)
    
    def _add_control(self, control_type, name, x, y, width, height, names=(), values=()):
        """ Add new control according to passed values. """
        dialog_model = self._dialog.getModel()
        control_model = dialog_model.createInstance(
            control_type if control_type.startswith("com.sun.star.awt") else "com.sun.star.awt.UnoControl" + control_type + "Model")
        dialog_model.insertByName(name, control_model)
        
        control_model.setPropertyValues(
            ("Height", "PositionX", "PositionY", "Width"), 
            (height, x, y, width))
        if len(names) > 0:
            #print(names, values)
            control_model.setPropertyValues(names, values)
        return self._dialog.getControl(name)
    
    def _get_control(self, name):
        """ Get control from dialog. """
        return self._dialog.getControl(name)
        
    
    def get_label(self, label_name):
        """ Get label data from a control. """
        return self._dialog.getControl(label_name).getModel().Label
    
    def set_label(self, label_name, label):
        """ Set label data from a control. """
        self._dialog.getControl(label_name).getModel().Label = label
    
    def get_text(self, edit_name):
        """ Get text data from an edit control. """
        return self._dialog.getControl(edit_name).getModel().Text
    
    def set_text(self, edit_name, text):
        """ Set text data from an edit control. """
        self._dialog.getControl(edit_name).getModel().Text = text
    
    def get_radio_state(self, radio_name):
        """ Get state of radio button. """
        return self._dialog.getControl(radio_name).getModel().State
    
    def get_check_state(self, check_name):
        return self._dialog.getControl(check_name).getModel().State == 1
    
    def set_check_state(self, check_name, state):
        self._dialog.getControl(check_name).getModel().State = 1 if state else 0
    
    def set_focus(self, name):
        """ Set focus on a control. """
        self._dialog.getControl(name).setFocus()
    
    def _set_visible(self, state):
        """ Change visibility of the dialog. """
        if self._disposed: return
        self._dialog.setVisible(state)
    
    def set_state(self, name, state):
        if self._disposed: return
        self._dialog.getControl(name).setEnable(state)
    
    def set_visible(self, name, state):
        self._dialog.getControl(name).setVisible(state)
    
 
class ModalDialog(DialogBase):
    """ Modal dialog base class. """
    def __init__(self, ctx, frame, dialog_url="", string_resource={}):
        DialogBase.__init__(self, ctx, frame, string_resource)
        
        self._result = 0
    
    def execute(self):
        """ Show dialog and returns result of the pushed button. """
        if not self._dialog: return None
        self._set_visible(True)
        self._result = self._dialog.execute()
        return self._result
    
    def dispose_dialog(self):
        """ Destroy dialog. """
        self._dispose()
 
 
from com.sun.star.awt import XTopWindowListener
from com.sun.star.frame import XFrameActionListener
 
from com.sun.star.frame.FrameAction import COMPONENT_DETACHING as FA_COMPONENT_DETACHING
 
 
class ModelessDialog(DialogBase):
    """ Modeless dialog base class. """
    
    class TopWindowListener(unohelper.Base, XTopWindowListener):
        def __init__(self, parent):
            self._parent = parent
        def disposing(self, ev):
            self._parent = None
        # XTopWindowListener
        def windowOpened(self, ev): pass
        def windowClosing(self, ev):
            """ Helps to close if X button pushed on the title bar of the dialog. """
            self._parent.close()
        def windowClosed(self, ev): pass
        def windowMinimized(self, ev): pass
        def windowNormalized(self, ev): pass
        def windowActivated(self, ev): pass
        def windowDeactivated(self, ev): pass
    
    class FrameActionListener(unohelper.Base, XFrameActionListener):
        def __init__(self, parent):
            self._parent = parent
        def disposing(self, ev):
            self._parent = None
        # XFrameActionListener
        def frameAction(self, ev):
            """ Close dialog if the component of the frame has been changed. """
            if ev.Action == FA_COMPONENT_DETACHING:
                self._parent._remove_frame_action_listener(self)
                self._parent.close()
    
    def __init__(self, ctx, frame, dialog_url="", string_resource={}):
        DialogBase.__init__(self, ctx, frame, dialog_url, string_resource={})
        self._frame_listener = None
        self._topwindow_listener = None
    
    def _set_visible(self, state):
        """ Change visibility of the dialog. """
        if self._disposed: return
        self._dialog.setVisible(state)
    
    def _create_dialog(self, x=0, y=0, width=0, height=0, title=""):
        """ Creates dialog with additional listeners. """
        DialogBase._create_dialog(self, x, y, width, height, title)
        
        window_listener = self.TopWindowListener(self)
        frame_listener = self.FrameActionListener(self)
        self._dialog.addTopWindowListener(window_listener)
        self._frame.addFrameActionListener(frame_listener)
        self._topwindow_listener = window_listener
        self._frame_listener = frame_listener
    
    def show(self):
        """ Show dialog. """
        self._set_visible(True)
    
    def hide(self):
        """ Hide dialog. """
        self._set_visible(False)
    
    def close(self):
        """ Close and destroy dialog. """
        self.hide()
        self._dispose()
    
    # XEventListener
    def disposing(self, ev):
        self._remove_frame_action_listener(self._frame_listener)
        DialogBase.disposing(self, ev)
        self._frame_listener = None
        self._topwindow_listener = None
    
    def _remove_frame_action_listener(self, listener):
        self._frame.removeFrameActionListener(listener)
 
 
from com.sun.star.beans.PropertyState import DIRECT_VALUE as PS_DIRECT_VALUE
from com.sun.star.beans import PropertyValue
 
from com.sun.star.sheet import XRangeSelectionListener
 
import re
 
class RangeSelectionDialog(ModelessDialog):
    """ Modeless dialog with range selection functions. """
    
    range_selection_locked = False
    
    RANGE_SELECTION_ACTION_PREFIX = "__RANGE_SELECTION__"
    
    class RangeSelectionOptions(object):
        """ Used to pass range selection options,  
            see css.sheet.RangeSelectionArguments service.
        """
        def __init__(self, edit_name, title, initial_value="", close_on_release=True, single_mode=False):
            self.edit_name = edit_name
            self.title = title
            self.initial_value = initial_value
            self.close_on_release = close_on_release
            self.single_mode = single_mode
    
    
    class RangeSelectionListenerImpl(unohelper.Base, XRangeSelectionListener):
        """ Tracks range selection changes. """
        def __init__(self, parent, edit_name):
            self._parent = parent
            self.edit_name = edit_name
        # XEventListener
        def disposing(self, ev):
            self._parent = None
        # XRangeSelectionListener
        def done(self, ev):
            desc = ev.RangeDescriptor
            desc = self._parse_descriptor(desc)
            self._parent.end_range_selection(self, self.edit_name, desc)
        
        def aborted(self, ev):
            self._parent.end_range_selection(self)
        
        def _parse_descriptor(self, desc):
            return desc.replace("$", "")
    
    class RangeSelectionButtonListenerImpl(unohelper.Base, XActionListener):
        def __init__(self, parent):
            self._parent = parent
        def disposing(self, ev):
            self._parent = None
        # XActionListener
        def actionPerformed(self, ev):
            command = ev.ActionCommand
            if command.startswith(self._parent.RANGE_SELECTION_ACTION_PREFIX):
                name = command[len(self._parent.RANGE_SELECTION_ACTION_PREFIX):]
                options = self._parent._range_selection_map[name]
                initial_value = options.initial_value if options.initial_value else self._parent.get_text(options.edit_name)
                self._parent.start_range_selection(options, initial_value)
    
    SMALL_BUTTON_WIDTH = 14
    SMALL_BUTTON_HEIGHT = 16
    
    def __init__(self, ctx, frame, dialog_url="", string_resource={}, range_selection_image_url=""):
        """ Initialize dialog. 
        @param ctx component context
        @param frame keeps spreadsheet document
        @param dialog_url dialog url to show if you want to use a dialog 
                            created by the dialog editor.
        @param string_resource translation table for the runtime dialog creation
        @param range_selection_image_url image url for range selection buttons which created at runtime
        """
        ModelessDialog.__init__(self, ctx, frame, dialog_url, string_resource)
        
        self.range_selection_image_url = range_selection_image_url
        self._controller = frame.getController()
        self._range_selection_listener = None
        self._range_selection_button_listener = self.RangeSelectionButtonListenerImpl(self)
        
        self._range_selection_map = {}
    
    def start_range_selection(self, options, initial_value=""):
        """ Starting range selection mode. """
        edit_name = options.edit_name
        if initial_value is None:
            initial_value = self.get_text(edit_name) if edit_name else ""
        args = self._get_range_selection_arguments(options, initial_value)
        #initial_value, title, close_on_release, single_mode)
        
        self.hide()
        
        self._range_selection_listener = self.RangeSelectionListenerImpl(self, edit_name)
        self._controller.addRangeSelectionListener(self._range_selection_listener)
        self._controller.startRangeSelection(args)
        
        self.range_selection_locked = True
    
    def end_range_selection(self, listener, edit_name=None, desc=""):
        """ Ends range selection mode. """
        if edit_name:
            self.set_text(edit_name, desc)
        #print("remove listener")
        self.show() # should be shown before the listener is removed
        self._controller.removeRangeSelectionListener(listener)
        self.range_selection_locked = False
        self._range_selection_listener = None
    
    def _get_range_selection_arguments(self, options, initial_value=""):
        """ Convert options to css.beans.PropertyValues """
        d = PS_DIRECT_VALUE
        return (
            PropertyValue("InitialValue", 0, initial_value, d), 
            PropertyValue("Title", 0, options.title, d), 
            PropertyValue("CloseOnMouseRelease", 0, options.close_on_release, d), 
            PropertyValue("SingleCellMode", 0, options.single_mode, d)
            )
    
    def _add_range_selection_button(self, name, x, y, width=SMALL_BUTTON_WIDTH, height=SMALL_BUTTON_HEIGHT):
        """ Adds button allows to start range selection. """
        return self._add_control("Button", name, x, y, width, height, 
            ("ImageURL", ), (self.range_selection_image_url,))
    
    def _add_edit(self, name, x, y, width, height, names=(), values=()):
        """ Adds edit control. """
        self._add_control("Edit", name, x, y, width, height, names, values)
    
    def _add_range_selection_set(self, edit_name, button_name, x, y, width, height, label_name="", 
        title="", title_from_label=True, initial_value="", on_mouse_release=True, 
        single_mode=False, named_range=False, range_names=None):
        """ Adds a set of range selection which has an edit control 
            shows range address and a push button to start range selection.
        
        """
        V_SHIFT = (self.SMALL_BUTTON_HEIGHT - height) / 2
        if named_range and isinstance(range_names, tuple):
            range_edit = self._add_control("ComboBox", edit_name, x, y, width - self.SMALL_BUTTON_WIDTH - 1, height, 
            ("Dropdown", "LineCount"), (True, 7))
            range_edit.addItems(range_names, 0)
        else:
            self._add_edit(edit_name, x, y, width - self.SMALL_BUTTON_WIDTH - 1, height)
        btn = self._add_range_selection_button(button_name, x + width - self.SMALL_BUTTON_WIDTH, y - V_SHIFT)
        
        self._range_selection_registration(btn, edit_name, button_name, label_name, 
            title, title_from_label, initial_value, on_mouse_release, single_mode)
    
    def _register_range_selection_set(self, edit_name, button_name, label_name="", 
        title="", title_from_label=True, initial_value="", on_mouse_release=True, 
        single_mode=False, named_range=False, range_names=None):
        """ Register a set of controls allows to do range selection. """
        
        if named_range and isinstance(range_names, tuple):
            range_edit = self._get_control(edit_name)
            range_edit.addItems(range_names, 0)
        
        self._range_selection_registration(edit_name, button_name, label_name, 
            title, title_from_label, initial_value, on_mouse_release, single_mode)
        
    
    def _range_selection_registration(self, edit_name, button_name, label_name, 
        title, title_from_label, initial_value, on_mouse_release, single_mode):
        btn = self._get_control(button_name)
        btn.addActionListener(self._range_selection_button_listener)
        btn.setActionCommand(self.RANGE_SELECTION_ACTION_PREFIX + button_name)
        
        if title_from_label and label_name:
            title = self._get_label_as_title(label_name)
        options = self.RangeSelectionOptions(edit_name, title, initial_value, on_mouse_release, single_mode)
        
        self._range_selection_map[button_name] = options
    
    def _get_label_as_title(self, name):
        label = self.get_label(name)
        label_b = re.sub("\(~\S\)", "", label).replace("~", "")
        return label_b
    
    def messagebox(self, message, title, message_type, buttons):
        parent = self._frame.getContainerWindow()
        toolkit = parent.getToolkit()
        
        msgbox = toolkit.createMessageBox(parent, Rectangle(), message_type, buttons, title, message)
        return msgbox.execute()
    
    def message(self, message, title=""):
        """ Shows information. """
        return self.messagebox(message, title, "messbox", 1)
    
    def warning(self, message, title=""):
        """ Shows warning message. """
        return self.messagebox(message, title, "warningbox", 1)
    
    def error(self, message, title=""):
        """ Shows error message. """
        return self.messagebox(message, title, "errorbox", 1)
    
# from dialog import RangeSelectionDialog
 
from com.sun.star.awt import XActionListener
 
class CustomSelectionDialog(RangeSelectionDialog, XActionListener):
    """ Example dialog uses RangeSelectionDialog class.
        
        The dialog created by the dialog editor and it has the following controls:
            buttonOk: OK button to start processing
            buttonCancel: Cancel button
            labelRange1: label for the range selection 1
            editRange1: first edit field shows in the single selection mode
            buttonRange1: button to start range selection
            labelRange2: label for the range selection 2
            editRange1: the combo box shows range selected and combo box keeps list of named ranges
            buttonRange2: to start
    """
    
    def __init__(self, manager, ctx, frame):
        dialog_url = ""  # "vnd.sun.star.script:Standard.Dialog6?location=application"
        self.image_url = "file:///home/asuka/Documents/images/se.png"
        RangeSelectionDialog.__init__(self, ctx, frame, dialog_url=dialog_url)
        self.manager = manager
    
    def _initialize_dialog(self):
        """ set listeners on the dialog controls. """
        dialog = self._dialog
        gc = self._get_control
        
        # set listeners to buttons
        buttonOk = gc("buttonOk")
        buttonCancel = gc("buttonCancel")
        buttonOk.addActionListener(self)
        buttonCancel.addActionListener(self)
        buttonOk.setActionCommand("ok")
        buttonCancel.setActionCommand("cancel")
        
        gc("buttonRange1").getModel().ImageURL = self.image_url
        gc("buttonRange2").getModel().ImageURL = self.image_url
        
        range_names = self.manager.get_named_range_names()
        
        """ editRange1 is used in single selection mode
        """
        self._register_range_selection_set("editRange1", "buttonRange1", label_name="labelRange1", 
            title="", title_from_label=True, initial_value="", 
            on_mouse_release=True, single_mode=True)
        
        """ editRange2 is the combo box and it shows named range
        """
        self._register_range_selection_set("editRange2", "buttonRange2", label_name="labelRange2", 
            title="", title_from_label=True, initial_value="", 
            on_mouse_release=True, single_mode=False, named_range=True, range_names=range_names)
        
    
    def actionPerformed(self, ev):
        cmd = ev.ActionCommand
        if cmd == "ok":
            # if you want to check the values on the dialog, try in here
            # validation...
            # call hide method if you want to hide the dialog before any processing
            self.hide()
            self.manager.do_something()
            self.close()
        elif cmd == "cancel":
            # close and destroy the dialog, dispose method of the dialog is called also
            self.close()
    
    def get_selection1(self):
        """ Get first selected. """
        return self.get_text("editRange1")
    
    def get_selection2(self):
        """ Get secondary selected. """
        return self.get_text("editRange2")
    
    def start(self):
        """ Show dialog. """
        if not self._dialog:
            self._create_dialog() # DialogBase method
            self._initialize_dialog()
        self.show()
 
 
class Manager(object):
    """ Simple test application with the dialog. """
    def __init__(self, ctx, frame):
        self._dialog = CustomSelectionDialog(self, ctx, frame)
        self._frame = frame
        self.ctx = ctx
    
    def start(self):
        """ Start this. """
        self._dialog.start()
    
    def do_something(self):
        """ OK button pushed on the dialog. """
        message = "Range 1: %s\nRange 2: %s" % (
            self._dialog.get_selection1(), self._dialog.get_selection2())
        
        self._dialog.message(message)
    
    def get_named_range_names(self):
        """ list of named range is taken from the dialog. """
        doc = self._get_document()
        return doc.NamedRanges.getElementNames()
    
    def _get_document(self):
        return self._frame.getController().getModel()
 
 
# def modeless_dialog_test(*args):
#     """ Just an example to use range selection dialog. """
#     ctx = XSCRIPTCONTEXT.getComponentContext()
#     doc = XSCRIPTCONTEXT.getDocument()

def main(ctx, smgr):
    
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
#     prop = PropertyValue(Name="Hidden", Value=True)
#     desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))  # バックグラウンドでWriterのドキュメントを開く。
    doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
    
    frame = doc.getCurrentController().getFrame()
    # should be check the document type is spreadsheet document
    
    manager = Manager(ctx, frame)
    manager.start()
    
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
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