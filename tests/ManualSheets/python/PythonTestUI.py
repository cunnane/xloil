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
# Qt pane creation in a try/except in case qtpy is not installed.
#

try:
    # You must import `xloil.gui.qtpy` before `qtpy`, this allows xlOil to create
    # a thread to manage the Qt GUI.  *All* interaction with the Qt GUI except emitting 
    # signals must be done on the GUI thread or Qt _will abort_.  Use `Qt_thread.submit(...)`
    # to send jobs to Qt's thread.
    import xloil.gui.qtpy
    
    from qtpy.QtWidgets import QLabel, QWidget, QHBoxLayout, QPushButton, QProgressBar
    from qtpy.QtCore import Signal, Qt
    
    class OurQtPane(QWidget):
        
        _progress = Signal(int)
        
        def set_progress(self, x: int):
            # Use a signal to send the progress: this is thread safe
            self._progress.emit(x)
            
        def __init__(self):
            super().__init__() # Must call this or Qt will crash
            
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

# Unlike Qt, it's not (I think) common to derive the from a tkinter object.
# Instead, we derive from `TkThreadTaskPane`, which derives from `CustomTaskPane`

class OurTkPane(TkThreadTaskPane):
        
    def __init__(self):
        super().__init__() # Important!
        
        import tkinter as tk
        
        top_level = self.top_level

        btn = tk.Button(top_level, text="This is a Button", fg='blue')
        btn.place(x=20, y=50)
        
        from tkinter import ttk
        self._progress_bar = ttk.Progressbar(top_level, length=200, mode='determinate')
        self._progress_bar.place(x=20, y=100)
    
    @Tk_thread
    def set_progress(self, x: int):
        self._progress_bar['value'] = x

    # Define this method to capture the docking position change event
    def on_docked(self):
        xlo.log(f"Tk frame docking position: {self.position:}", level='info')


# Create a wxPython task pane, but wrap in a try..except in case wx is not installed
try:
    from xloil.gui.wx import wx_thread
    import wx

    class OurWxPane(wx.Frame):
        def __init__(self):
            # ensure the parent's __init__ is called
            super().__init__(None, title='Hello')

            # create a panel in the frame
            pnl = wx.Panel(self)

            # put some text with a larger bold font on it
            st = wx.StaticText(pnl, label="Hello World!")
            font = st.GetFont()
            font.PointSize += 10
            font = font.Bold()
            st.SetFont(font)

            self._gauge = wx.Gauge(pnl)

            # and create a sizer to manage the layout of child widgets
            sizer = wx.BoxSizer(wx.VERTICAL)
            sizer.Add(st, wx.SizerFlags().Border(wx.TOP|wx.LEFT, 25))
            sizer.Add(self._gauge, wx.SizerFlags().Border(wx.TOP|wx.LEFT, 25))
            pnl.SetSizer(sizer)

        @wx_thread
        def set_progress(self, x: int):
            self._gauge.SetValue(x)

except ImportError:
    class OurWxPane:
        ...

_PENDING_PANES = dict()

_PANE_NAMES = { 
    'Tk': "MyTkPane", 
    'Qt': "MyQtPane",
    'wx': "MyWxPane"
}

# We define a function to create a task pane using Tk or Qt. We first check 
# that the pane has not already been created, then construct a instance of the
# pane class, then attach it to the ExcelGUI object created later.

async def make_task_pane(toolkit):

    global _excelgui
    
    pane_name = _PANE_NAMES[toolkit]

    key = (pane_name, xlo.app().windows.active.name)

    pane = xlo.gui.find_task_pane(pane_name)
    if pane is not None:
        xlo.log(f"Found pane: {key}")
        pane.visible = True
        return

    # Since we are using async, the open pane button may have been 
    # clicked more than once before the pane got a chance to create
    pane_future = _PENDING_PANES.get(key, None)
    if pane_future is not None:
        return

    # attach_pane can accept a CustomTaskPane instance, or a QWidget
    # instance or an awaitable to one of those things. You can also
    # pass a QWidget type which xlOil wrap in a QtThreadTaskPane and
    # create in the correct thread.
    #
    # (We could just pass `OurTkPane` rather than using Tk_thread, this
    # is just to demonstrate passing an awaitable)
    if toolkit == 'Tk':
        future = _excelgui.attach_pane_async(
            name=pane_name,
            pane=Tk_thread().submit_async(OurTkPane))
    elif toolkit == 'Qt':
        future = _excelgui.attach_pane_async(
            name=pane_name, 
            pane=OurQtPane)
    elif toolkit == 'wx':
        future = _excelgui.attach_pane_async(
            name=pane_name, 
            pane=OurWxPane)
    else:
        raise Exception()

    _PENDING_PANES[key] = future
        
    pane = await future
    
    del _PENDING_PANES[key]

    pane.visible = True

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
    "buttonQt": "Qt",
    "buttonWx": "wx"
} 

def get_button_label(ctrl, *args):
    # Ribbon callback to determine button label text
    return f"Open {_BUTTON_MAP[ctrl.id]}"

async def press_open_pane_button(ctrl):
    
    toolkit = _BUTTON_MAP[ctrl.id]
    
    xlo.log(f"Open {toolkit} Pressed")
    
    await make_task_pane(toolkit)

    
#
# The combo box in the ribbon xml has the value 33, 66 or 99. We send 
# this as the progress % to the progress bar in our task panes (if they 
# have been created)
# 
def combo_change(ctrl, value):
    
    qt_pane = xlo.gui.find_task_pane(_PANE_NAMES['Qt'])
    if qt_pane:
        qt_pane.widget.set_progress(int(value))
        
    tk_pane = xlo.gui.find_task_pane(_PANE_NAMES['Tk'])
    if tk_pane:
        tk_pane.set_progress(int(value))

    wx_pane = xlo.gui.find_task_pane(_PANE_NAMES['wx'])
    if wx_pane:
        wx_pane.frame.set_progress(int(value))

    return "NotSupposedToReturnHere" # check this doesn't cause an error

async def press_open_console_button_tk(ctrl):

    def sesame(root):
        from xloil.gui.tkinter import TkConsole
        import tkinter
        import code

        top_level = tkinter.Toplevel(root)
        console = TkConsole(top_level, code.interact,
            fg='white', bg='black', font='Consolas', insertbackground='red')
        console.pack(expand=True, fill=tkinter.BOTH)
        console.bind("<<CommandDone>>", lambda e: top_level.destroy())

        top_level.deiconify()

    from xloil.gui.tkinter import Tk_thread
    await Tk_thread().submit_async(sesame, Tk_thread().root)

async def press_open_console_button_qt(ctrl):

    def sesame():
        from xloil.gui.qt_console import create_qtconsole_inprocess
        console = create_qtconsole_inprocess()
        console.show()
        return console

    from xloil.gui.qtpy import Qt_thread
    await Qt_thread().submit_async(sesame)

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
                        <button id="buttonWx" getLabel="getButtonLabel" getImage="buttonImg" size="large" onAction="pressOpenPane" />
                        <button id="tkConsole" label="Tk Console" size="large" onAction="pressOpenConsoleTk" />
                        <button id="qtConsole" label="Qt Console" size="large" onAction="pressOpenConsoleQt" />
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
        'pressOpenConsoleTk': press_open_console_button_tk,
        'pressOpenConsoleQt': press_open_console_button_qt,
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

