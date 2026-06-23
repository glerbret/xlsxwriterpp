/*
 * utility - Utility functions for Xlsxwriter++.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

/// #ifdef USE_FMEMOPEN
/// #define _POSIX_C_SOURCE 200809L
/// #endif

#include "xwpp/utility.h"

#include "xlsxwriterpp.h"

#include <algorithm>
#include <chrono>

#include <iostream>

/// #ifdef USE_DTOA_LIBRARY
/// #include "xlsxwriter/third_party/emyg_dtoa.h"
/// #endif

namespace xwpp
{

std::string col_to_name(col_num_t col_num, bool absolute)
{
  std::string col_name;

  // Change from 0 index to 1 index.
  col_num++;

  // Convert the column number to a string in reverse order.
  while(col_num != 0)
  {
    // Get the remainder in base 26.
    int remainder = col_num % 26;
    if(remainder == 0)
    {
      remainder = 26;
    }

    // Convert the remainder value to a character.
    col_name.push_back(static_cast<char>('A' + remainder - 1));

    // Get the next order of magnitude.
    col_num = (col_num - 1) / 26;
  }

  if(absolute)
  {
    col_name.push_back('$');
  }

  std::ranges::reverse(col_name);

  return col_name;
}

std::string rowcol_to_cell(row_num_t row, col_num_t col)
{
  return col_to_name(col, false) + std::to_string(row + 1);
}

std::string rowcol_to_cell_abs(row_num_t row, col_num_t col, bool abs_row, bool abs_col)
{
  std::string cell_name = col_to_name(col, abs_col);
  if(abs_row)
  {
    cell_name += '$';
  }
  cell_name += std::to_string(row + 1);

  return cell_name;
}

std::string rowcol_to_range(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col)
{
  // Add the first cell to the range.
  std::string range = rowcol_to_cell(first_row, first_col);

  // If the start and end cells are the same just return a single cell.
  if(first_row == last_row && first_col == last_col)
  {
    return range;
  }

  // Add the range separator.
  range.push_back(':');

  // Add the first cell to the range.
  range += rowcol_to_cell(last_row, last_col);

  return range;
}

std::string rowcol_to_range_abs(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col)
{
  // Add the first cell to the range.
  std::string range = rowcol_to_cell_abs(first_row, first_col, 1, 1);

  // If the start and end cells are the same just return a single cell.
  if (first_row == last_row && first_col == last_col)
    return range;

  // Add the range separator.
  range += ':';

  // Add the first cell to the range.
  range += rowcol_to_cell_abs(last_row, last_col, 1, 1);

  return range;
}

std::string rowcol_to_formula_abs(const std::string& sheetname, row_num_t first_row, col_num_t first_col,
                                  row_num_t last_row, col_num_t last_col)
{
  std::string formula = quote_sheetname(sheetname);

  // Add the range separator.
  formula += '!';

  // Add the first cell to the range.
  formula += rowcol_to_cell_abs(first_row, first_col, true, true);

  // If the start and end cells are the same just return a single cell.
  if(first_row == last_row && first_col == last_col)
  {
    return formula;
  }

  // Add the range separator.
  formula += ':';

  // Add the first cell to the range.
  formula += rowcol_to_cell_abs(last_row, last_col, true, true);

  return formula;
}

// TODO string
row_num_t name_to_row(const char* row_str)
{
  row_num_t row_num = 0;

  if(!row_str)
  {
    return row_num;
  }

  // Skip the column letters and absolute symbol of the A1 cell.
  while(*row_str && !isdigit((unsigned char)*row_str))
  {
    row_str++;
  }

  // Convert the row part of the A1 cell to a number.
  if(*row_str)
  {
    row_num = atoi(row_str);
  }

  if(row_num)
  {
    row_num--;
  }

  return row_num;
}

uint32_t name_to_row_2(const char* row_str)
{
  if(!row_str)
  {
    return 0;
  }

  // Find the : separator in the range.
  while(*row_str && *row_str != ':')
  {
    row_str++;
  }

  if(*row_str)
  {
    return name_to_row(++row_str);
  }
  else
  {
    return 0;
  }
}

// TODO String
col_num_t name_to_col(const char* col_str)
{
  col_num_t col_num = 0;

  if(!col_str)
  {
    return col_num;
  }

  // Convert leading column letters of A1 cell. Ignore absolute $ marker.
  while(*col_str && (isupper((unsigned char)*col_str) || *col_str == '$'))
  {
    if(*col_str != '$')
    {
      col_num = (col_num * 26) + (*col_str - 'A' + 1);
    }
    col_str++;
  }

  if(col_num)
  {
    col_num--;
  }

  return col_num;
}

uint16_t name_to_col_2(const char* col_str)
{
  if(!col_str)
  {
    return 0;
  }

  // Find the : separator in the range.
  while(*col_str && *col_str != ':')
  {
    col_str++;
  }

  if(*col_str)
  {
    return name_to_col(++col_str);
  }
  else
  {
    return 0;
  }
}

/// lxw_error
/// lxw_datetime_validate(lxw_datetime *datetime)
/// {
///     if (!datetime)
///         return LXW_ERROR_DATETIME_VALIDATION;

/*
 * Excel uses the year 1900 as the default epoch but it uses 1899-12-31 as
 * the 0 date and internally we use the 0-0-0 date for time only values.
 */
///     if (datetime->year < 1900 &&
///         !(datetime->year == 0 &&
///           datetime->month == 0 && datetime->day == 0) &&
///         !(datetime->year == 1899 &&
///           datetime->month == 12 && datetime->day == 31)) {

///         LXW_WARN_FORMAT1("lxw_datetime_validate(): invalid year: %d. "
///                          "Valid range is 1900-9999.", datetime->year);

///         return LXW_ERROR_DATETIME_VALIDATION;
///     }

///     if (datetime->year > 9999) {
///         LXW_WARN_FORMAT1("lxw_datetime_validate(): invalid year: %d. "
///                          "Valid range is 1900-9999.", datetime->year);
///         return LXW_ERROR_DATETIME_VALIDATION;
///     }

///     if (datetime->year != 0) {
///         if (datetime->month < 1 || datetime->month > 12) {
///             LXW_WARN_FORMAT1("lxw_datetime_validate(): invalid month: %d. "
///                              "Valid range is 1-12.", datetime->month);
///             return LXW_ERROR_DATETIME_VALIDATION;
///         }

///         if (datetime->day < 1 || datetime->day > 31) {
///             LXW_WARN_FORMAT1("lxw_datetime_validate(): invalid day: %d. "
///                              "Valid range is 1-31.", datetime->day);
///             return LXW_ERROR_DATETIME_VALIDATION;
///         }
///     }

///     if (datetime->hour < 0 || datetime->hour > 23) {
///         LXW_WARN_FORMAT1("lxw_datetime_validate(): invalid hour: %d. "
///                          "Valid range is 0-23.", datetime->hour);
///         return LXW_ERROR_DATETIME_VALIDATION;
///     }

///     if (datetime->min < 0 || datetime->min > 59) {
///         LXW_WARN_FORMAT1("lxw_datetime_validate(): invalid minute: %d. "
///                          "Valid range is 0-59.", datetime->min);
///         return LXW_ERROR_DATETIME_VALIDATION;
///     }

///     if (datetime->sec < 0.0 || datetime->sec >= 60.0) {
///         LXW_WARN_FORMAT1("lxw_datetime_validate(): invalid seconds: %.3f. "
///                          "Valid range is 0.0-59.999.", datetime->sec);
///         return LXW_ERROR_DATETIME_VALIDATION;
///     }

///     return LXW_NO_ERROR;
/// }

double datetime_to_excel_date_with_epoch(const std::chrono::system_clock::time_point& datetime, bool use_1904_epoch)
{
  const auto dp = floor<std::chrono::days>(datetime);
  const std::chrono::year_month_day ymd{dp};
  const std::chrono::hh_mm_ss time{std::chrono::floor<std::chrono::milliseconds>(datetime - dp)};
  const int year   = static_cast<int>(ymd.year());
  const int month  = static_cast<unsigned int>(ymd.month());
  const int day    = static_cast<unsigned int>(ymd.day());
  const int hour   = time.hours().count();
  const int min    = time.minutes().count();
  const double sec = time.seconds().count() + time.subseconds().count() / 1000.0;
  const int epoch  = use_1904_epoch ? 1904 : 1900;
  const int offset = use_1904_epoch ? 4 : 0;
  const int norm   = 300;
  // Set month days and check for leap year.
  int mdays[]      = {0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
  int leap         = 0;
  int days         = 0;
  int i;

  // Convert the Excel seconds to a fraction of the seconds in 24 hours.
  const double seconds = (hour * 60 * 60 + min * 60 + sec) / (24 * 60 * 60.0);

  // Special cases for Excel dates in the 1900 epoch.
  if(!use_1904_epoch)
  {
    // Excel 1900 epoch.
    if(year == 1899 && month == 12 && day == 31)
    {
      return seconds;
    }

    // Excel 1900 epoch.
    if(year == 1900 && month == 1 && day == 0)
    {
      return seconds;
    }

    // Excel false leapday
    if(year == 1900 && month == 2 && day == 29)
    {
      return 60 + seconds;
    }
  }

  /* We calculate the date by calculating the number of days since the
     epoch and adjust for the number of leap days. We calculate the
     number of leap days by normalizing the year in relation to the
     epoch. Thus the year 2000 becomes 100 for 4-year and 100-year
     leapdays and 400 for 400-year leapdays. */
  const int range = year - epoch;

  if(year % 4 == 0 && (year % 100 > 0 || year % 400 == 0))
  {
    leap     = 1;
    mdays[2] = 29;
  }

  // Calculate the serial date by accumulating the number of days
  //  since the epoch.

  // Add days for previous months.
  for(i = 0; i < month; i++)
  {
    days += mdays[i];
  }
  // Add days for current month.
  days += day;
  // Add days for all previous years.
  days += range * 365;
  // Add 4 year leapdays.
  days += range / 4;
  // Remove 100 year leapdays.
  days -= (range + offset) / 100;
  // Add 400 year leapdays.
  days += (range + offset + norm) / 400;
  // Remove leap days already counted.
  days -= leap;

  // Adjust for Excel erroneously treating 1900 as a leap year.
  if(!use_1904_epoch && days > 59)
  {
    days++;
  }

  return days + seconds;
}

/// double
/// lxw_datetime_to_excel_datetime(lxw_datetime *datetime)
/// {
///     return lxw_datetime_to_excel_date_with_epoch(datetime, LXW_FALSE);
/// }

/*
 * Convert a unix datetime (1970/01/01 epoch) to an Excel serial date, with a
 * 1900 epoch.
 */
/// double
/// lxw_unixtime_to_excel_date(int64_t unixtime)
/// {
///     return lxw_unixtime_to_excel_date_with_epoch(unixtime, LXW_FALSE);
/// }

// Convert a unix datetime (1970/01/01 epoch) to an Excel serial date, with a
// 1900 or 1904 epoch.
double unixtime_to_excel_date_with_epoch(int64_t unixtime, bool use_1904_epoch)
{
  double excel_datetime = 0.0;
  double epoch          = use_1904_epoch ? 24107.0 : 25568.0;

  excel_datetime = epoch + (unixtime / (24 * 60 * 60.0));

  if(!use_1904_epoch && excel_datetime >= 60.0)
  {
    excel_datetime = excel_datetime + 1.0;
  }

  return excel_datetime;
}
// Create a quoted version of the worksheet name, or return an unmodified
// copy if it doesn't required quoting.
std::string quote_sheetname(std::string_view sheetname)
{
  // Check if the sheetname contains any characters that require it
  // to be quoted. Also check for single quotes within the string.
  if(std::ranges::all_of(sheetname, [](char c) { return isalnum(c) != 0 || c == '_' || c == '.'; }))
  {
    return std::string{sheetname};
  }

  // Add single quotes to the start and end of the string.
  std::string quoted_name = "'";
  for(auto c: sheetname)
  {
    quoted_name.push_back(c);
    // Escape single quotes in name
    if(c == '\'')
    {
      quoted_name.push_back('\'');
    }
  }
  quoted_name.push_back('\'');
  return quoted_name;
}

/// FILE *
/// lxw_tmpfile(const char *tmpdir)
/// {
/// #ifndef USE_STANDARD_TMPFILE
///     return tmpfileplus(tmpdir, NULL, NULL, 0);
/// #else
///     (void) tmpdir;
///     return tmpfile();
/// #endif
/// }

/**
 * Return a memory-backed file if supported, otherwise a temporary one
 */
/// FILE *
/// lxw_get_filehandle(char **buf, size_t *size, const char *tmpdir)
/// {
///     static size_t s;
///     if (!size)
///         size = &s;
///     *buf = NULL;
///     *size = 0;
/// #ifdef USE_FMEMOPEN
///     (void) tmpdir;
///     return open_memstream(buf, size);
/// #else
///     return lxw_tmpfile(tmpdir);
/// #endif
/// }

/*
 * Use third party function to handle sprintf of doubles for locale portable
 * code.
 */
/// #ifdef USE_DTOA_LIBRARY
/// int
/// lxw_sprintf_dbl(char *data, double number)
/// {
///     emyg_dtoa(number, data);
///     return 0;
/// }
/// #endif

/// const char * lxw_version(void)
/// {
///     return LXW_VERSION;
/// }

/// uint16_t lxw_version_id(void)
/// {
///     return LXW_VERSION_ID;
/// }

/*
 * Hash a worksheet password. Based on the algorithm in ECMA-376-4:2016,
 * Office Open XML File Formats - Transitional Migration Features,
 * Additional attributes for workbookProtection element (Part 1, §18.2.29).
 */
uint16_t hash_password(const std::string& password)
{
  if(password.empty())
  {
    return 0;
  }

  uint16_t hash = 0;
  for(auto it = std::crbegin(password); it != std::crend(password); ++it)
  {
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    hash ^= *it & 0xFF;
  }

  hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
  hash ^= password.size();
  hash ^= 0xCE4B;

  return hash;
}

/* Make a simple portable version of fopen() for Windows. */

/// FILE *
/// lxw_fopen(const char *filename, const char *mode)
/// {
/// return fopen(filename, mode);
/// }

}
