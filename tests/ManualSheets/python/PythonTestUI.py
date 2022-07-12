# If not first, gives a "specified module cannot be found" error - ?
from PIL import Image

import xloil as xlo
from xloil.pandas import PDFrame
import datetime as dt
import asyncio

#---------------------------------
# GUI: Creating Custom Task Panes
#---------------------------------

# You *must* import `xloil.gui.pyqt5` before `PyQt5`, this allows xlOil to create
# a thread to manage the Qt GUI.  *All* interaction with the Qt GUI must be done
# on the GUI thread or Qt _will abort_.  Use `QtThread.submit(...)`

try:
    import xloil.gui.pyqt5

    _PANE_NAME="MyPane"

    from PyQt5.QtWidgets import QLabel, QWidget, QHBoxLayout, QPushButton, QProgressBar
    from PyQt5.QtCore import pyqtSignal, Qt
    
    class MyTaskPane(QWidget):
        progress = pyqtSignal(int)
    
        def __init__(self):
            super().__init__()
            progress_bar = QProgressBar(self)
            progress_bar.setGeometry(200, 80, 250, 20)
            self.progress.connect(progress_bar.setValue, Qt.QueuedConnection)
        
            label = QLabel("Hello from Qt")
        
            layout = QHBoxLayout()
        
            layout.addWidget(label)

            layout.addWidget(progress_bar)
        
            self.setLayout(layout)
            
            
    async def make_task_pane():
        global _excelgui # Will be created in time...
        return await _excelgui.create_task_pane(
            name=_PANE_NAME, creator=MyTaskPane)
            
except ImportError:
    raise
    
#----------------------
# GUI: Creating Ribbons
#----------------------
#
# xlOil is able to create a ribbon entry for a workbook which is automatically 
# removed when the workbook is closed and the associated workbook module is 
# unloaded.  To create the ribbon XML use an editor such as Office RibbonX Editor:
# https://github.com/fernandreu/office-ribbonx-editor
#
 
def get_icon_path():
    # Gets the path to an icon file to demonstrate PIL image handling
    return os.path.join(os.path.dirname(xloil.linked_workbook()), 'icon.bmp')
    
def button_label(ctrl, *args):
    return "Open Task Pane"
 
def button_image(ctrl):
    import os
    im = Image.open(get_icon_path())
    return im

#
# GUI callbacks declared async will be executed in the addin's 
# event loop. Other callbacks are executed in Excel's main thread.
# Async callbacks cannot return values.
# 
async def pressOpenPane(ctrl):
    
    xlo.log("Button Pressed")
    
    pane = await make_task_pane()
    pane.visible = True
    
def combo_change(ctrl, value):
    
    # The combo box has the value 33, 66 or 99. We send this as the progress % 
    # to the progress bar in our task pane (if it has been created)
    pane = xlo.find_task_pane(title=_PANE_NAME)
    if pane:
        xlo.log(f"Combo: {value} sent to progress bar")
        pane.widget.progress.emit(int(value))
    return "NotSupposedToReturnHere" # check this doesn't cause an error

#
# We construct the ExcelGUI (actually a handle to a COM addin) using XML to describe 
# the ribbon and a map from callbacks referred to in the XML to actual python functions
#
_excelgui = xlo.create_gui(ribbon=r'''
    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
        <ribbon>
            <tabs>
                <tab id="customTab" label="xloPyTest" insertAfterMso="TabHome">
                    <group idMso="GroupClipboard" />
                    <group idMso="GroupFont" />
                    <group id="customGroup" label="MyButtons">
                        <button id="pyButt1" getLabel="buttonLabel" getImage="buttonImg" size="large" onAction="pressOpenPane" />
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
    func_names={
        'pressOpenPane': pressOpenPane,
        'comboChange': combo_change,
        'buttonLabel': button_label,
        'buttonImg': button_image
    }).result()
    
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
        im = Image.open(get_icon_path())
        return im
    
    
    # Normally we use a return converter as an annotation like `-> ReturnImage` but 
    # if we want to dynamically pass arguments to the converter we can call it 
    # directly as below
    @xlo.func(macro=True)
    def pyTestPicSized(width:float, height:float, fitCell: bool=False):
        from PIL import Image
        import os
        im = Image.open(get_icon_path())
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

