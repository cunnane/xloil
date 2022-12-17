#pragma once
#include <stdint.h>

namespace xloil
{
  /// <summary>
  /// Max string length for an A1-style cell address. The largest address
  /// is "AAA1000000:ZZZ1000001<null>"
  /// </summary>
  constexpr uint16_t XL_CELL_ADDRESS_A1_MAX_LEN = 3 + 7 + 1 + 3 + 7 + 1;
  /// <summary>
  /// Max string length for an RC-style cell address
  /// </summary>
  constexpr uint16_t XL_CELL_ADDRESS_RC_MAX_LEN = 29;
  /// <summary>
  /// Max string length for a sheet name
  /// </summary>
  constexpr uint16_t XL_SHEET_NAME_MAX_LEN = 31;
  /// <summary>
  /// Max filename length. Used to be 216, but limit was raised
  /// in May 2020 Office 365. Length of 260 is imposed by filesystem.
  /// </summary>
  constexpr uint16_t XL_FILENAME_MAX_LEN = 260;
  /// <summary>
  /// Max string length for an A1-style full address including workbook name
  /// </summary>
  constexpr uint16_t XL_FULL_ADDRESS_A1_MAX_LEN = XL_FILENAME_MAX_LEN + XL_SHEET_NAME_MAX_LEN + XL_CELL_ADDRESS_A1_MAX_LEN;
  /// <summary>
  /// Max string length for an RC-style full address including workbook name
  /// </summary>
  constexpr uint16_t XL_FULL_ADDRESS_RC_MAX_LEN = XL_FILENAME_MAX_LEN + XL_SHEET_NAME_MAX_LEN + XL_CELL_ADDRESS_RC_MAX_LEN;
  /// <summary>
  /// Max string length for a (pascal) string in an ExcelObj
  /// </summary>
  constexpr uint16_t XL_STRING_MAX_LEN = 32767;
  /// <summary>
  /// Max number of rows on a sheet
  /// </summary>
  constexpr uint32_t XL_MAX_ROWS = 1048576;
  /// <summary>
  /// Max number of columns on a sheet
  /// </summary>
  constexpr uint16_t XL_MAX_COLS = 16384;
  /// <summary>
  /// Max number of args for a user-defined function
  /// </summary>
  constexpr uint16_t XL_MAX_UDF_ARGS = 255;
  /// <summary>
  /// Max number of args for a VBA function
  /// </summary>
  constexpr uint16_t XL_MAX_VBA_FUNCTION_ARGS = 60;
  /// <summary>
  /// Max number of args for a VBA function
  /// </summary>
  constexpr uint16_t XL_ARG_HELP_STRING_MAX_LENGTH = 255;
}