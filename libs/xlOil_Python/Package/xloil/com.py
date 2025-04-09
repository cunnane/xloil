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
    and switching to manual calculation mode.  Which of these changes are applied can
    be controlled by parameters to the constructor - the default is to apply all of them.
    Previous settings are restored when the context scope closes.

    Parameters
    ----------
    events: bool (default False)
        if False, pauses Excel's event model (Application.EnableEvents in VBA)

    calculation: bool (default False)
        if False, sets the calculation mode to manual (Application.Calculation in VBA)

    screen_updating: bool
        if False, disables screen updating in Excel (Application.ScreenUpdating in VBA)

    alerts: bool
        if False, disables alerts in Excel (Application.DisplayAlerts in VBA)

    """
    _calc_mode = None
    _screen_updating = None
    _events = None
    _alerts = None

    def __enter__(self, events=False, calculation=False, screen_updating=False, alerts=False):
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

        if not alerts:
            self._alerts = app.DisplayAlerts
            app.DisplayAlerts = False

        return self

    def __exit__(self, type, value, traceback):

        app = _core.app()

        if self._events is not None:
            app.EnableEvents = self._events

        if self._calc_mode is not None:
            app.Calculation = self._calc_mode

        if self._screen_updating is not None:
            app.ScreenUpdating = self._screen_updating

        if self._alerts is not None:
            app.DisplayAlerts = self._alerts


def fix_name_errors(workbook):
    """
        Marks #NAME! errors in current workbook for recalculation. Can be used
        when missing addin functions have caused #NAME! errors - these errors cannot
        simply be resolved by full recalc (Ctrl-Alt-F9).

        May not be performant in large workbooks with many errors.
    """
    for ws in workbook.worksheets:
        errors = ws.used_range.special_cells("formulas", "errors")
        if errors is None:
            continue
        for cell in errors:
            if cell.Value2 == _core.CellError.NAME.value:
                cell.Formula = cell.Formula
