/*
 * utility - Utility functions for Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/utility.h"

#include "xlsxwriterpp.h"
#include "xwpp/exception.h"

#include <algorithm>
#include <chrono>
#include <format>
#include <ranges>
#include <utility>
#include <vector>

namespace xwpp
{

std::string version()
{
  return XWPP_VERSION;
}

// Create a quoted version of the worksheet name, or return an unmodified
// copy if it doesn't required quoting.
std::string quote_sheetname(std::string_view sheetname)
{
  // Don't quote the sheetname if it is already quoted.
  if(sheetname[0] == '\'')
  {
    return std::string{sheetname};
  }

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
    // Escape single quotes in name.
    if(c == '\'')
    {
      quoted_name.push_back('\'');
    }
  }
  quoted_name.push_back('\'');
  return quoted_name;
}

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

std::string rowcol_to_cell(row_num_t row_num, col_num_t col_num)
{
  return std::format("{}{}", col_to_name(col_num, false), (row_num + 1));
}

std::string rowcol_to_cell_abs(row_num_t row_num, col_num_t col_num, bool abs_row, bool abs_col)
{
  std::string cell_name = col_to_name(col_num, abs_col);
  if(abs_row)
  {
    cell_name += '$';
  }
  cell_name += std::to_string(row_num + 1);

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
  std::string range = rowcol_to_cell_abs(first_row, first_col, true, true);

  // If the start and end cells are the same just return a single cell.
  if(first_row == last_row && first_col == last_col)
  {
    return range;
  }

  // Add the range separator.
  range += ':';

  // Add the first cell to the range.
  range += rowcol_to_cell_abs(last_row, last_col, true, true);

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

row_num_t name_to_row(std::string_view row_str)
{
  row_num_t row_num = 0;

  if(row_str.empty())
  {
    return row_num;
  }

  // Skip the column letters and absolute symbol of the A1 cell.
  // Convert the row part of the A1 cell to a number.
  if(const auto found = row_str.find_first_of("0123456789"); found != std::string_view::npos)
  {
    row_num = std::stoul(std::string{row_str.substr(found)});
  }

  if(row_num != 0)
  {
    row_num--;
  }

  return row_num;
}

uint32_t name_to_row_2(std::string_view row_str)
{
  if(row_str.empty())
  {
    return 0;
  }

  // Find the : separator in the range.
  const auto found = row_str.find_first_of(':');
  if(found != std::string_view::npos)
  {
    return name_to_row(row_str.substr(found + 1));
  }
  else
  {
    return 0;
  }
}

col_num_t name_to_col(std::string_view col_str)
{
  col_num_t col_num = 0;

  if(col_str.empty())
  {
    return col_num;
  }

  for(const auto c: col_str)
  {
    if(isupper(c) != 0)
    {
      col_num = static_cast<col_num_t>((col_num * 26) + (c - 'A' + 1));
    }
    else if(c == '$')
    {
      // NOP
    }
    else
    {
      // Stop the reading
      break;
    }
  }

  if(col_num != 0)
  {
    col_num--;
  }

  return col_num;
}

uint16_t name_to_col_2(std::string_view col_str)
{
  if(col_str.empty())
  {
    return 0;
  }

  // Find the : separator in the range.
  const auto found = col_str.find_first_of(':');
  if(found != std::string_view::npos)
  {
    return name_to_col(col_str.substr(found + 1));
  }
  else
  {
    return 0;
  }
}

std::string dup_formula(const std::string& formula)
{
  if(formula[0] == '=')
  {
    return formula.substr(1);
  }
  else
  {
    return formula;
  }
}

double pixels_to_width(double pixels)
{
  const double max_digit_width = 7.0;
  const double padding         = 5.0;
  double width                 = 0.;

  if(pixels == DEF_COL_WIDTH_PIXELS)
  {
    width = DEF_COL_WIDTH;
  }
  else if(pixels <= 12.0)
  {
    width = pixels / (max_digit_width + padding);
  }
  else
  {
    width = (pixels - padding) / max_digit_width;
  }

  return width;
}

double pixels_to_height(double pixels)
{
  if(pixels == DEF_ROW_HEIGHT_PIXELS)
  {
    return DEF_ROW_HEIGHT;
  }
  else
  {
    return pixels * 0.75;
  }
}

datetime_t to_datetime(const std::chrono::system_clock::time_point& datetime)
{
  const auto date = std::chrono::floor<std::chrono::days>(datetime);
  const std::chrono::year_month_day ymd{date};
  const std::chrono::hh_mm_ss time{std::chrono::floor<std::chrono::milliseconds>(datetime - date)};

  const auto year  = static_cast<int>(ymd.year());
  const auto month = static_cast<int>(static_cast<unsigned int>(ymd.month()));
  const auto day   = static_cast<int>(static_cast<unsigned int>(ymd.day()));

  // time_point set to epoch date (1970-01-01) are time only
  if(year == 1970 && month == 1 && day == 1)
  {
    return datetime_t{.year_  = 0,
                      .month_ = 0,
                      .day_   = 0,
                      .hour_  = static_cast<int>(time.hours().count()),
                      .min_   = static_cast<int>(time.minutes().count()),
                      .sec_   = static_cast<double>(time.seconds().count()) +
                              (static_cast<double>(time.subseconds().count()) / 1000.0)};
  }

  return datetime_t{.year_  = year,
                    .month_ = month,
                    .day_   = day,
                    .hour_  = static_cast<int>(time.hours().count()),
                    .min_   = static_cast<int>(time.minutes().count()),
                    .sec_   = static_cast<double>(time.seconds().count()) +
                            (static_cast<double>(time.subseconds().count()) / 1000.0)};
}

double datetime_to_excel_date_with_epoch(const std::chrono::system_clock::time_point& datetime, bool use_1904_epoch)
{
  return datetime_to_excel_date_with_epoch(to_datetime(datetime), use_1904_epoch);
}

double datetime_to_excel_datetime(const std::chrono::system_clock::time_point& datetime)
{
  return datetime_to_excel_date_with_epoch(datetime, false);
}

double datetime_to_excel_date_with_epoch(const datetime_t& datetime, bool use_1904_epoch)
{
  int year         = datetime.year_;
  int month        = datetime.month_;
  int day          = datetime.day_;
  const int hour   = datetime.hour_;
  const int min    = datetime.min_;
  const double sec = datetime.sec_;
  const int epoch  = use_1904_epoch ? 1904 : 1900;
  const int offset = use_1904_epoch ? 4 : 0;
  const int norm   = 300;
  // Set month days and check for leap year.
  std::vector<int> mdays{0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
  int leap = 0;
  int days = 0;

  datetime_validate(datetime);

  // For times without dates set the default date for the epoch.
  if(year == 0)
  {
    if(!use_1904_epoch)
    {
      year  = 1899;
      month = 12;
      day   = 31;
    }
    else
    {
      year  = 1904;
      month = 1;
      day   = 1;
    }
  }

  // Convert the Excel seconds to a fraction of the seconds in 24 hours.
  const double seconds = ((hour * 60 * 60) + (min * 60) + sec) / (24 * 60 * 60.0);

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

  // We calculate the date by calculating the number of days since the
  // epoch and adjust for the number of leap days. We calculate the
  // number of leap days by normalizing the year in relation to the
  // epoch. Thus the year 2000 becomes 100 for 4-year and 100-year
  // leapdays and 400 for 400-year leapdays.
  const int range = year - epoch;

  if(year % 4 == 0 && (year % 100 > 0 || year % 400 == 0))
  {
    leap     = 1;
    mdays[2] = 29;
  }

  // Calculate the serial date by accumulating the number of days
  // since the epoch.

  // Add days for previous months.
  for(size_t i = 0; std::cmp_less(i, month) && i < mdays.size(); i++)
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

double datetime_to_excel_datetime(const datetime_t& datetime)
{
  return datetime_to_excel_date_with_epoch(datetime, false);
}

void datetime_validate(const datetime_t& datetime)
{
  // Excel uses the year 1900 as the default epoch but it uses 1899-12-31 as
  // the 0 date and internally we use the 0-0-0 date for time only values.
  if(datetime.year_ < 1900 && (datetime.year_ != 0 || datetime.month_ != 0 || datetime.day_ != 0) &&
     (datetime.year_ != 1899 || datetime.month_ != 12 || datetime.day_ != 31))
  {
    throw xwpp_exception_t(
      std::format("datetime_validate(): invalid year: {}. Valid range is 1900-9999.", datetime.year_));
  }

  if(datetime.year_ > 9999)
  {
    throw xwpp_exception_t(
      std::format("datetime_validate(): invalid year: {}. Valid range is 1900-9999.", datetime.year_));
  }

  if(datetime.year_ != 0)
  {
    if(datetime.month_ < 1 || datetime.month_ > 12)
    {
      throw xwpp_exception_t(
        std::format("datetime_validate(): invalid month: {}. Valid range is 1-12.", datetime.month_));
    }

    if(datetime.day_ < 1 || datetime.day_ > 31)
    {
      throw xwpp_exception_t(std::format("datetime_validate(): invalid day: {}. Valid range is 1-31.", datetime.day_));
    }
  }

  if(datetime.hour_ < 0 || datetime.hour_ > 23)
  {
    throw xwpp_exception_t(std::format("datetime_validate(): invalid hour: {}. Valid range is 0-23.", datetime.hour_));
  }

  if(datetime.min_ < 0 || datetime.min_ > 59)
  {
    throw xwpp_exception_t(std::format("datetime_validate(): invalid minute: {}. Valid range is 0-59.", datetime.min_));
  }

  if(datetime.sec_ < 0.0 || datetime.sec_ >= 60.0)
  {
    throw xwpp_exception_t(
      std::format("datetime_validate(): invalid seconds: {}. Valid range is 0.0-59.999.", datetime.sec_));
  }
}

double unixtime_to_excel_date_with_epoch(int64_t unixtime, bool use_1904_epoch)
{
  const double epoch    = use_1904_epoch ? 24107.0 : 25568.0;
  double excel_datetime = epoch + (static_cast<double>(unixtime) / (24 * 60 * 60.0));

  if(!use_1904_epoch && excel_datetime >= 60.0)
  {
    excel_datetime = excel_datetime + 1.0;
  }

  return excel_datetime;
}

double unixtime_to_excel_date(int64_t unixtime)
{
  return unixtime_to_excel_date_with_epoch(unixtime, false);
}

uint16_t hash_password(const std::string& password)
{
  if(password.empty())
  {
    return 0;
  }

  uint16_t hash = 0;
  for(const unsigned char c: std::ranges::reverse_view(password))
  {
    hash = (static_cast<uint16_t>(hash >> 14U) & 0x01U) | (static_cast<uint16_t>(hash << 1U) & 0x7fffU);
    hash ^= c & 0xFFU;
  }

  hash = (static_cast<uint16_t>(hash >> 14U) & 0x01U) | (static_cast<uint16_t>(hash << 1U) & 0x7fffU);
  hash ^= password.size();
  hash ^= 0xCE4BU;

  return hash;
}

std::string to_lower(const std::string& str)
{
  std::string lower_str;

  std::ranges::transform(str, std::back_inserter(lower_str), [](unsigned char c) { return std::tolower(c); });

  return lower_str;
}

}
