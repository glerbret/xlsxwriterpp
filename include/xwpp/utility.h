/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @file
 *
 * @brief Utility functions for `Xlsxwriter++`.
 */

#ifndef XWPP_UTILITY_H
#define XWPP_UTILITY_H

#include "common.h"

#include <string>
#include <string_view>

/// @cond
namespace xwpp
{
/// @endcond

/**
 * @brief Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *  worksheet.write_string(CELL("A1"), "Foo");
 *
 *  //Same as:
 *  worksheet.write_string(0, 0, "Foo");
 * @endcode
 *
 * @note This macro shouldn't be used in performance critical situations since it
 * expands to two function calls.
 *
 * @todo Replace by a constexpr function or overload of function using `%col_num`
 * and `%row_num`.
 */
#define CELL(cell) xwpp::name_to_row(cell), xwpp::name_to_col(cell)

/**
 * @brief Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *  worksheet.set_column(COLS("B:D"), 20);
 *
 *  // Same as:
 *  worksheet_set_column(1, 3, 20);
 * @endcode
 *
 * @todo Replace by a constexpr function or overload of function using `%col_num`
 * and `%row_num`.
 */
#define COLS(cols) xwpp::name_to_col(cols), xwpp::name_to_col_2(cols)

/**
 * @brief Convert an Excel `A1:B2` range into a `(first_row, first_col,
 * last_row, last_col)` sequence.
 *
 * Convert an Excel `A1:B2` range into a `(first_row, first_col, last_row,
 * last_col)` sequence.
 *
 * This is a little syntactic shortcut to help with worksheet layout.
 *
 * @code
 *  worksheet.print_area(RANGE("A1:K42"));
 *
 *  // Same as:
 *  worksheet.print_area(0, 0, 41, 10);
 * @endcode
 *
 * @todo Replace by a constexpr function or overload of function using `%col_num`
 * and `%row_num`.
 */
#define RANGE(range)                                                                                                   \
  xwpp::name_to_row(range), xwpp::name_to_col(range), xwpp::name_to_row_2(range), xwpp::name_to_col_2(range)

/**
 * @brief Retrieve the library version.
 *
 * @return The "X.Y.Z" version string.
 *
 * Get the library version as a "X.Y.Z" version string
 *
 * @code
 *  std::cout << std::format("Xlsxwriter++ version = {}\n", xwpp::version());
 * @endcode
 */
std::string version();

// Create a quoted version of the worksheet name
[[nodiscard]] std::string quote_sheetname(std::string_view sheetname);

[[nodiscard]] std::string col_to_name(col_num_t col_num, bool absolute);
[[nodiscard]] std::string rowcol_to_cell(row_num_t row, col_num_t col);
[[nodiscard]] std::string rowcol_to_cell_abs(row_num_t row, col_num_t col, bool abs_row, bool abs_col);
[[nodiscard]] std::string rowcol_to_range(row_num_t first_row, col_num_t first_col, row_num_t last_row,
                                          col_num_t last_col);
[[nodiscard]] std::string rowcol_to_range_abs(row_num_t first_row, col_num_t first_col, row_num_t last_row,
                                              col_num_t last_col);
std::string rowcol_to_formula_abs(const std::string& sheetname, row_num_t first_row, col_num_t first_col,
                                  row_num_t last_row, col_num_t last_col);

// TODO String
uint32_t name_to_row(const char* row_str);
uint32_t name_to_row_2(const char* row_str);
uint16_t name_to_col(const char* col_str);
uint16_t name_to_col_2(const char* col_str);

std::string dup_formula(const std::string& formula);

/**
 * @brief Converts a `%system_clock::time_point` to an Excel datetime number with
 * 1900/1904 epoch.
 *
 * @param datetime A `%system_clock::time_point`.
 * @param use_1904_epoch A flag to indicate whether to use the 1904 epoch (true)
 *  or the 1900 epoch (false).
 *
 * This function is similar to `datetime_to_excel_datetime()` but it allows
 * you to specify whether to use the 1900 or 1904 epoch.
 *
 * @see The `workbook_t::use_1904_epoch()` function.
 *
 * @todo Add again the overload using "lxw_datetime".
 * @todo Add Other overloads.
 * @todo Add Note about range of std::chrono::system_clock::time_point(up to 2062).
 */
double datetime_to_excel_date_with_epoch(const std::chrono::system_clock::time_point& datetime, bool use_1904_epoch);

/**
 * @brief Converts a `%system_clock::time_point` to an Excel datetime number.
 *
 * @param datetime A `%system_clock::time_point`.
 *
 * @return A double representing an Excel datetime.
 *
 * The `%datetime_to_excel_datetime()` function converts a datetime in
 * `%system_clock::time_point` to an Excel datetime number:
 *
 * @code
 *  std::chrono::sys_days{2013y / std::chrono::February / 28d} + 12h + 0min + 0s;
 *
 *  double excel_datetime = xwpp::datetime_to_excel_date(datetime);
 * @endcode
 *
 * @see @ref working_with_dates for more details on the Excel datetime format.
 *
 * @todo Add again the overload using "lxw_datetime" (and don't forget document "working with datetime").
 * @todo Add Other overloads.
 * @todo Add Note about range of std::chrono::system_clock::time_point(up to 2062).
 */
double datetime_to_excel_datetime(const std::chrono::system_clock::time_point& datetime);

// /**
//  * @brief Validate a #lxw_datetime struct.
//  *
//  * Validates a #lxw_datetime struct to ensure its fields are within acceptable
//  * ranges for Excel dates and times.
//  *
//  * The members of the #lxw_datetime struct and the range of their values are:
//  *
//  * Member   | Value
//  * -------- | -----------
//  * year     | 1900 - 9999
//  * month    | 1 - 12
//  * day      | 1 - 31
//  * hour     | 0 - 23
//  * min      | 0 - 59
//  * sec      | 0 - 59.999
//  *
//  * @param datetime A pointer to a #lxw_datetime struct.
//  *
//  * @return A #lxw_error code. Either #LXW_NO_ERROR or
//  *         #LXW_ERROR_DATETIME_VALIDATION if a field is out of range.
//  */
// TODO Set again lxw_datetime
// lxw_error lxw_datetime_validate(lxw_datetime *datetime);

/**
 * @brief Converts a unix datetime to an Excel datetime number with 1900/1904
 * epoch.
 *
 * @param unixtime Unix time (seconds since 1970-01-01)
 * @param use_1904_epoch A flag to indicate whether to use the 1904 epoch (true)
 *        or the 1900 epoch (false).
 *
 * This function is similar to `unixtime_to_excel_date()` but it allows
 * you to specify whether to use the 1900 or 1904 epoch. See also the
 * `workbook_t::use_1904_epoch()` function.
 */
double unixtime_to_excel_date_with_epoch(int64_t unixtime, bool use_1904_epoch);

/**
 * @brief Converts a unix datetime to an Excel datetime number.
 *
 * @param unixtime Unix time (seconds since 1970-01-01)
 *
 * @return A double representing an Excel datetime.
 *
 * The `%unixtime_to_excel_date()` function converts a unix datetime to
 * an Excel datetime number:
 *
 * @code
 *  double excel_datetime = xwpp::unixtime_to_excel_date(946684800);
 * @endcode
 *
 * @see @ref working_with_dates for more details.
 */
double unixtime_to_excel_date(int64_t unixtime);

// Hash a worksheet password. Based on the algorithm in ECMA-376-4:2016,
// Office Open XML File Formats - Transitional Migration Features,
// Additional attributes for workbookProtection element (Part 1, §18.2.29).
uint16_t hash_password(const std::string& password);

std::string to_lower(const std::string& str);

/// @cond
}
/// @endcond

#endif
