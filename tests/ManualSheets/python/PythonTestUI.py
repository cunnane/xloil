# If PIL is not first, gives a "specified module cannot be found" error - ?
try:
    from PIL import Image
except ImportError:
    pass
    
import xloil as xlo
from xloil.pandas import PDFrame
import datetime as dt
import asyncio
import inspect

#---------------------------------
# GUI: Creating Custom Task Panes
#---------------------------------
#
# We demonstrate how task panes can be created using Qt and Tk. We wrap the 
# Qt pane creation in a try/except in case pyqt5 is not installed.
#

try:
    # You must import `xloil.gui.pyqt5` before `PyQt5`, this allows xlOil to create
    # a thread to manage the Qt GUI.  *All* interaction with the Qt GUI except emitting 
    # signals must be done on the GUI thread or Qt _will abort_.  Use `Qt_thread.submit(...)`
    # to send jobs to Qt's thread.
    import xloil.gui.pyqt5
    
    from PyQt5.QtWidgets import QLabel, QWidget, QHBoxLayout, QPushButton, QProgressBar
    from PyQt5.QtCore import pyqtSignal, Qt
    
    class OurQtPane(QWidget):
        
        _progress = pyqtSignal(int)
        
        def set_progress(self, x: int):
            # Use a signal to send the progress: this is thread safe
            self._progress.emit(x)
            
        def __init__(self):
            super().__init__()
            progress_bar = QProgressBar(self)
            progress_bar.setGeometry(200, 80, 250, 20)
            self._progress.connect(progress_bar.setValue, Qt.QueuedConnection)
        
            label = QLabel("Hello from Qt")
        
            layout = QHBoxLayout()
        
            layout.addWidget(label)

            layout.addWidget(progress_bar)
        
            self.setLayout(layout)
            
except ImportError:

    class OurQtPane:
        ...
   
# Like Qt, xlOil's tkinter module must be imported before using the toolkit
# to allow xlOil to create the *Tk_thread* and the tkinter root object. All
# interactions with tkinter must take place via *Tk_thread*.
from xloil.gui.tkinter import TkThreadTaskPane, Tk_thread

# Unlike Qt, it's not common to derive the from a tkinter object.
# Instead, we derive from `TkThreadTaskPane`, which derives from `CustomTaskPane`
# We provide a `draw` method to create a `tkinter.Toplevel` which comprises of
# the pane's contents
class OurTkPane(TkThreadTaskPane):
    
    name = 'TkPane'
    
    @Tk_thread
    def set_progress(self, x: int):
        self._progress_bar['value'] = x
       
    def draw(self):
    
        import tkinter as tk
        
        top_level = tk.Toplevel()

        btn = tk.Button(top_level, text="This is a Button", fg='blue')
        btn.place(x=20, y=50)
        
        from tkinter import ttk
        self._progress_bar = ttk.Progressbar(top_level, length=200, mode='determinate')
        self._progress_bar.place(x=20, y=100)
        
        return top_level

#
# We define a function to create a task pane using Tk or Qt. We first check 
# that the pane has not already been created, then construct a instance of the
# pane class, then attach it to the ExcelGUI object created later.
#

_PANES = {}

async def make_task_pane(toolkit):

    global _excelgui
    
    # We need to be careful in case the open pane button was clicked
    # more than once before the pane got a chance to create
    pane = _PANES.get(toolkit, None)
    if pane is not None:
        if inspect.isfuture(pane):
            return await pane
        else:
            return pane
    
    # Note the Qt creator just returns the type OurQtPane: it's allowable 
    # to pass a QWidget type (or an instance) into `attach_pane`. Passing
    # the type is safer so xlOil can ensure it's created on the Qt thread.
    # xlOil will wrap the widget in a QtThreadTaskPane.

    if toolkit == 'Tk':
        future = _excelgui.attach_pane_async(OurTkPane())
    elif toolkit == 'Qt':
        future = _excelgui.attach_pane_async(name="MyQtPane", pane=OurQtPane)
    else:
        raise Exception()

    _PANES[toolkit] = future
        
    pane = await future
    
    _PANES[toolkit] = pane
    
    return pane
     
#----------------------
# GUI: Creating Ribbons
#----------------------
#
# xlOil is able to create a ribbon entry for a workbook which is automatically 
# removed when the workbook is closed and the associated workbook module is 
# unloaded.  To create the ribbon XML use an editor such as Office RibbonX Editor:
# https://github.com/fernandreu/office-ribbonx-editor
#
# GUI callbacks declared async will be executed in the addin's 
# event loop. Other callbacks are executed in Excel's main thread.
# Async callbacks cannot return values.
# 

def _get_icon_path():
    import os
    # Gets the path to an icon file to demonstrate PIL image handling
    return os.path.join(os.path.dirname(xloil.linked_workbook()), 'icon.bmp')
    

def button_image(ctrl):
    # Ribbon callback to determine the button's icon (see ribbon xml)
    im = Image.open(_get_icon_path())
    return im

# Maps button ids in the ribbon xml below to GUI toolkit names
_BUTTON_MAP = { 
    "buttonTk": "Tk", 
    "buttonQt": "Qt"
} 

def get_button_label(ctrl, *args):
    # Ribbon callback to determine button label text
    return f"Open {_BUTTON_MAP[ctrl.id]}"

async def press_open_pane_button(ctrl):
    
    toolkit = _BUTTON_MAP[ctrl.id]
    
    xlo.log(f"Open {toolkit} Pressed")
    
    pane = await make_task_pane(toolkit)
    pane.visible = True
    
#
# The combo box in the ribbon xml has the value 33, 66 or 99. We send 
# this as the progress % to the progress bar in our task panes (if they 
# have been created)
# 
def combo_change(ctrl, value):
    
    qt_pane = _PANES.get('Qt', None)
    if qt_pane:
        qt_pane.contents.set_progress(int(value))
        
    tk_pane = _PANES.get('Tk', None)
    if tk_pane:
        tk_pane.set_progress(int(value))
      
    return "NotSupposedToReturnHere" # check this doesn't cause an error

#
# We construct the ExcelGUI (actually a handle to a COM addin) using XML to describe 
# the ribbon and a map from callbacks referred to in the XML to actual python functions
#
_excelgui = xlo.ExcelGUI(ribbon=r'''
    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
        <ribbon>
            <tabs>
                <tab id="customTab" label="xloPyTest" insertAfterMso="TabHome">
                    <group idMso="GroupClipboard" />
                    <group idMso="GroupFont" />
                    <group id="customGroup" label="MyButtons">
                        <button id="buttonTk" getLabel="getButtonLabel" getImage="buttonImg" size="large" onAction="pressOpenPane" />
                        <button id="buttonQt" getLabel="getButtonLabel" getImage="buttonImg" size="large" onAction="pressOpenPane" />
                        <comboBox id="comboBox" label="Combo Box" onChange="comboChange">
                         <item id="item1" label="33" />
                         <item id="item2" label="66" />
                         <item id="item3" label="99" />
                       </comboBox>
                    </group>
                </tab>
            </tabs>
        </ribbon>
    </customUI>
    ''', 
    funcmap={
        'pressOpenPane': press_open_pane_button,
        'comboChange': combo_change,
        'getButtonLabel': get_button_label,
        'buttonImg': button_image
    })
    
#-----------------------------------------
# Images: returning images from functions
#-----------------------------------------

# In case PIL is not installed, protect this section
try:

    # This import defines a return converter which allows us to return a PIL image
    import xloil.pillow
    import os
    from PIL import Image
    
    # The image return converter is registered, so we just need to return the PIL
    # image from an xlo.func. Returning an image requires macro=True permissions
    @xlo.func(macro=True)
    def pyTestPic():
        im = Image.open(_get_icon_path())
        return im
    
    
    # Normally we use a return converter as an annotation like `-> ReturnImage` but 
    # if we want to dynamically pass arguments to the converter we can call it 
    # directly as below
    @xlo.func(macro=True)
    def pyTestPicSized(width:float, height:float, fitCell: bool=False):
        from PIL import Image
        import os
        im = Image.open(_get_icon_path())
        if fitCell:
            return xlo.pillow.ReturnImage(size="cell")(im)
        else:
            return xlo.pillow.ReturnImage((width, height))(im)

except ImportError:
    pass

#-----------------------------------------
# Plots: returning matplotlib figures from functions
#-----------------------------------------

# In case matplotlib is not installed, protect this section
try:

    # This import defines a return converter for a matplotlib figure
    # It also imports matplotlib like this:
    #
    #   import matplotlib
    #   matplotlib.use('Agg')
    #   from matplotlib import pyplot
    # 
    # The order is important: the matplotlib backend must be switched
    # from the Qt default before pyplot is imported
    #
    
    import xloil.matplotlib
    from matplotlib import pyplot
    
    @xlo.func(macro=True)
    def pyTestPlot(x, y, width:float=4, height:float=4, **kwargs):
        fig = pyplot.figure()
        fig.set_size_inches(width, height)
        ax = fig.add_subplot(111)
        ax.plot(x, y, **kwargs)
        return fig
        
except ImportError:
    pass

