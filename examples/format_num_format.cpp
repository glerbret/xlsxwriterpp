/*
 * Example of writing some data with numeric formatting to a simple Excel file
 * using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 30);

  // Add some formats.
  xwpp::format_t* format01 = workbook.add_format();
  xwpp::format_t* format02 = workbook.add_format();
  xwpp::format_t* format03 = workbook.add_format();
  xwpp::format_t* format04 = workbook.add_format();
  xwpp::format_t* format05 = workbook.add_format();
  xwpp::format_t* format06 = workbook.add_format();
  xwpp::format_t* format07 = workbook.add_format();
  xwpp::format_t* format08 = workbook.add_format();
  xwpp::format_t* format09 = workbook.add_format();
  xwpp::format_t* format10 = workbook.add_format();
  xwpp::format_t* format11 = workbook.add_format();

  // Set some example number formats.
  format01->set_num_format("0.000");
  format02->set_num_format("#,##0");
  format03->set_num_format("#,##0.00");
  format04->set_num_format("0.00");
  format05->set_num_format("mm/dd/yy");
  format06->set_num_format("mmm d yyyy");
  format07->set_num_format("d mmmm yyyy");
  format08->set_num_format("dd/mm/yyyy hh:mm AM/PM");
  format09->set_num_format("0 \"dollar and\" .00 \"cents\"");

  worksheet.write_number(2, 0, 123);

  // Write data using the formats.
  worksheet.write_number(0, 0, 3.1415926);           // 3.1415926
  worksheet.write_number(1, 0, 3.1415926, format01); // 3.142
  worksheet.write_number(2, 0, 1234.56, format02);   // 1,235
  worksheet.write_number(3, 0, 1234.56, format03);   // 1,234.56
  worksheet.write_number(4, 0, 49.99, format04);     // 49.99
  worksheet.write_number(5, 0, 36892.521, format05); // 01/01/01
  worksheet.write_number(6, 0, 36892.521, format06); // Jan 1 2001
  worksheet.write_number(7, 0, 36892.521, format07); // 1 January 2001
  worksheet.write_number(8, 0, 36892.521, format08); // 01/01/2001 12:30 AM
  worksheet.write_number(9, 0, 1.87, format09);      // 1 dollar and .87 cents

  // Show limited conditional number formats.
  format10->set_num_format("[Green]General;[Red]-General;General");
  worksheet.write_number(10, 0, 123, format10); // > 0 Green
  worksheet.write_number(11, 0, -45, format10); // < 0 Red
  worksheet.write_number(12, 0, 0, format10);   // = 0 Default color

  // Format a Zip code.
  format11->set_num_format("00000");
  worksheet.write_number(13, 0, 1209, format11); // 01209

  workbook.save("format_num_format.xlsx");
}
