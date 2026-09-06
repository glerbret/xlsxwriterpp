/*
 * Example of writing a dates and time in Excel using a number with date
 * formatting. This demonstrates that dates and times in Excel are just
 * formatted real numbers.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // A number to display as a date.
  const double number{41333.5};

  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a format with date formatting.
  xwpp::format_t* format = workbook.add_format();
  format->set_num_format("mmm d yyyy hh:mm AM/PM");

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 22);

  // Write the number without formatting.
  worksheet.write_number(0, 0, number); // 41333.5

  /* Write the number with formatting. Note: the worksheet_write_datetime()
   * or worksheet_write_unixtime() functions are preferable for writing
   * dates and times. This is for demonstration purposes only.
   */
  worksheet.write_number(1, 0, number, format); // Feb 28 2013 12:00 PM

  workbook.save("date_and_times01.xlsx");
}
