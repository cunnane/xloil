from . import _core

# TODO: support comtypes
class _Win32ComConstants:
    def __getattr__(self, name):
        from win32com.client import constants
        return getattr(constants, name)


constants = _Win32ComConstants()
"""
    Contains the numeric value for enums used in the Excel.Application API.
    For example:

    ::

        from xloil import constants as xlc
        xloil.app().Calculation = xlc.xlCalculationManual
        xloil.app().Selection.PasteSpecial(Paste=xlc.xlPasteFormulas)

"""

class PauseExcel():
    """
    A context manager which pauses Excel by disabling events, turning off screen updating 
    and switching to manual calculation mode.  Which of these changees are applied can
    be controlled by parameters to the constructor - the default is to apply all of them.
    """
    _calc_mode = None
    _screen_updating = None
    _events = None

    def __enter__(self, events=False, calculation=False, screen_updating=False):
        app = _core.app()

        if not events:
            self._events = app.EnableEvents
            app.EnableEvents = False

        if not calculation:
            self._calc_mode = app.Calculation
            # The below constant equals constants.xlCalculationManual but
            # avoids the dependency on win32com
            app.Calculation = -4135

        if not screen_updating:
            self._screen_updating = app.ScreenUpdating
            app.ScreenUpdating = False

        return self

    def __exit__(self, type, value, traceback):

        app = _core.app()

        if self._events is not None:
            app.EnableEvents = self._events

        if self._calc_mode is not None:
            app.Calculation = self._calc_mode

        if self._screen_updating is not None:
            app.ScreenUpdating = self._screen_updating


def fix_name_errors(workbook):
    """
        Marks #NAME! errors in current workbook for recalculation. Can be used
        when missing addin functions have caused #NAME! errors - these errors cannot
        simply be resolved by full recalc (Ctrl-Alt-F9).

        May not be performant in large workbooks with many errors.
    """
    for ws in workbook.worksheets:
        try:
            # Unfortunately we have to drop into COM as we don't have an implementation
            # of SpecialCells in xloil
            for cell in ws.UsedRange.SpecialCells(constants.xlCellTypeFormulas, constants.xlErrors):
                if cell.Value2 == _core.CellError.NAME.value:
                    cell.Formula = cell.Formula
        except:
            # SpecialCells throws if it's empty so we catch and ignore
            pass
