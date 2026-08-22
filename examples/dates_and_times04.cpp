/*
 * Example of writing dates and times in Excel using different date formats.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <chrono>
#include <vector>

using namespace std::literals::chrono_literals;

int main()
{
  // A datetime to display.
  const auto datetime = std::chrono::sys_days{2013y / std::chrono::January / 23d} + 12h + 30min + 5s + 123ms;

  // Examples date and time formats. In the output file compare how changing
  // the format strings changes the appearance of the date.
  const std::vector<std::string> date_formats{
    "dd/mm/yy",
    "mm/dd/yy",
    "dd m yy",
    "d mm yy",
    "d mmm yy",
    "d mmmm yy",
    "d mmmm yyy",
    "d mmmm yyyy",
    "dd/mm/yy hh:mm",
    "dd/mm/yy hh:mm:ss",
    "dd/mm/yy hh:mm:ss.000",
    "hh:mm",
    "hh:mm:ss",
    "hh:mm:ss.000",
  };

  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a bold format.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Write the column headers.
  worksheet.write_string(0, 0, "Formatted date", bold);
  worksheet.write_string(0, 1, "Format", bold);

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 1, 22);

  // Write the same date and time using each of the above formats.
  for(xwpp::row_num_t row_num = 1; const auto& date_format: date_formats)
  {
    // Create a format for the date or time.
    xwpp::format_t* format = workbook.add_format();
    format->set_num_format(date_format);
    format->set_align(xwpp::format_alignments_t::HORIZONTAL_LEFT);

    // Write the datetime with each format.
    worksheet.write_datetime(row_num, 0, datetime, format);

    // Also write the format string for comparison.
    worksheet.write_string(row_num, 1, date_format);

    row_num++;
  }

  workbook.save("date_and_times04.xlsx");
}
