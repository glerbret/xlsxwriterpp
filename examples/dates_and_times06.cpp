/*
 * Example of writing dates and times in Excel using a datetime_t struct
 * and date formatting.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <ctime>

int main()
{
  // A datetime to display.
  const struct tm datetime = {.tm_hour = 12, .tm_mday = 28, .tm_mon = 1, .tm_year = 113};

  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a format with date formatting.
  xwpp::format_t* format = workbook.add_format();
  format->set_num_format("mmm d yyyy hh:mm AM/PM");

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 22);

  // Write the datetime without formatting.
  worksheet.write_datetime(0, 0, datetime); // 41333.5

  // Write the datetime with formatting.
  worksheet.write_datetime(1, 0, datetime, format); // Feb 28 2013 12:00 PM

  workbook.save("date_and_times06.xlsx");
}
