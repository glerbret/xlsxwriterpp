/*
 * Example of writing dates and times in Excel using a Unix datetime and date
 * formatting.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <ctime>

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a format with date formatting.
  xwpp::format_t* format = workbook.add_format();
  format->set_num_format("mmm d yyyy hh:mm AM/PM");

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 22);

  // Write some Unix datetimes with formatting.
  // 1970-01-01. The Unix epoch.
  worksheet.write_datetime(0, 0, static_cast<time_t>(0), format);

  // 2000-01-01.
  worksheet.write_datetime(1, 0, static_cast<time_t>(1577836800LL), format);

  // 1900-01-01.
  worksheet.write_datetime(2, 0, static_cast<time_t>(-2208988800LL), format);

  workbook.save("date_and_times04.xlsx");
}
