# LibreOffice Python Macro Collection on GUI

This repository contains a PyDev package.

#### Environment

- linuxBean　(Ubuntu) 14.04 
  
- LibreOffice 5.2.6.2

- Eclipse Neon.2 (4.6.2)

- PyDev 5.6.0

    - These macros also work on Windows 10.

## Macros

### Modeless Dialog Examples

#### Three methods for creating a modeless dialog: 

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/macro/modelessdialog2macro_createWindow.py">modelessdialog2macro_createWindow.py</a>

    - Created by using Toolkit method createWindow()

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/macro/modelessdialog2macro_taskcreator.py">modelessdialog2macro_taskcreator.py</a>

    - Created by using TaskCreator Service

    - This medhod only for creating a modeless dialog.

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/macro/modelessdialog2macro_unocontroldialog.py">modelessdialog2macro_unocontroldialog.py</a>

    - Created by using UnoControlDialog Service

####  A modeless dialog linking with a Writer document:

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/macro/modelessdialog3macro.py">modelessdialog3macro.py</a>

  - The text selected in a Writer document is reflected in the edit control on a modeless dialog.

### Splitter Examples

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/splitter/splitter_createWindow.py">splitter_createWindow.py</a>

  - Splitter on the window created by using Toolkit method createWindow()

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/splitter/splitter_unocontroldialog.py">splitter_unocontroldialog.py</a>

  - Splitter on the window created by using UnoControlDialog Service

### UnoControlDialog with various controls

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/unodialogsample.py">unodialogsample.py</a>

  - Currency fields

  - Dividing lines

  - Text fields

  - Time fields

  - Date fields

  - Fields of groups 

  - Pattern fields

  - Numerical fields

  - Checkboxes

  - Radio Buttons

  - List boxes

  - Combo-boxes (not work)

  - Formatted fields

  - Scrollbars

  - File selection fields 

  - Standard buttons

  - Hyperlink

  - Roadmap

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/unodialogsample2.py">unodialogsample2.py</a>

  - Roadmap

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/unomenu2.py">unomenu2.py</a>

  - PopupMenu

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/imagecontrol/imagecontrolsample.py">imagecontrolsample.py</a>

  - ImageControl

### UnoControlDialog with a subframe

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/dialogdocument.py">dialogdocument.py</a>

  - Loading a Writer document onto a component window of the subframe 

### A Sizeable Dialog

UnoControlDialog can not be sizeable.

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/imagecontrol/imagecontrolsample_sizeable3_forWin.py">imagecontrolsample_sizeable3_forWin.py</a>

  - The size of the ImageControl follows the size of the window created by using TaskCreator Service.

### Dialogs with a menu bar

You can not display a menu bar in UnoControlDialog.

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/unomenu_createWindow.py">unomenu_createWindow.py</a>

  - A menu bar in the window created by using Toolkit method createWindow()

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/unomenu_taskcreator.py">unomenu_taskcreator.py</a>

  - A menu bar in the window created by using using TaskCreator Service
  
###  File Selection Dialog Examples

- <a href="https://github.com/p--q/GUI/blob/master/GUI/src/systemdialog.py">systemdialog.py</a>

  - File selection dialogs with com.sun.star.ui.dialogs.TemplateDescription as a template are displayed in order.
  
     - FILEOPEN_SIMPLE, FILEOPEN_LINK_PREVIEW_IMAGE_TEMPLATE, FILEOPEN_PLAY, FILEOPEN_READONLY_VERSION, FILEOPEN_LINK_PREVIEW

     - FILEOPEN_PREVIEW, FILEOPEN_LINK_PLAY (LibreOffice 5.3 and higher only)

     - FILESAVE_SIMPLE, FILESAVE_AUTOEXTENSION_PASSWORD, FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS, FILESAVE_AUTOEXTENSION_SELECTION, FILESAVE_AUTOEXTENSION_TEMPLATE, FILESAVE_AUTOEXTENSION

### Release Notes

2018-1-2 version 1.0.2 Fixed a serious bug. Changed to `setattr(UnoObject, aPropertyName, aValue)` since it is no longer possible to use tuples in aValue of `setPropertyValue([in] string aPropertyName, [in] any aValue)` from LibreOffice 5.4.
