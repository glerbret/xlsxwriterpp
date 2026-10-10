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

  // Set some example number formats.
  const xwpp::format_t* format01 = workbook.format_builder().num_format("0.000").build();
  const xwpp::format_t* format02 = workbook.format_builder().num_format("#,##0").build();
  const xwpp::format_t* format03 = workbook.format_builder().num_format("#,##0.00").build();
  const xwpp::format_t* format04 = workbook.format_builder().num_format("0.00").build();
  const xwpp::format_t* format05 = workbook.format_builder().num_format("mm/dd/yy").build();
  const xwpp::format_t* format06 = workbook.format_builder().num_format("mmm d yyyy").build();
  const xwpp::format_t* format07 = workbook.format_builder().num_format("d mmmm yyyy").build();
  const xwpp::format_t* format08 = workbook.format_builder().num_format("dd/mm/yyyy hh:mm AM/PM").build();
  const xwpp::format_t* format09 = workbook.format_builder().num_format(R"(0 "dollar and" .00 "cents")").build();

  // Write data using the formats.
  worksheet.write(0, 0, 3.1415926);           // 3.1415926
  worksheet.write(1, 0, 3.1415926, format01); // 3.142
  worksheet.write(2, 0, 1234.56, format02);   // 1,235
  worksheet.write(3, 0, 1234.56, format03);   // 1,234.56
  worksheet.write(4, 0, 49.99, format04);     // 49.99
  worksheet.write(5, 0, 36892.521, format05); // 01/01/01
  worksheet.write(6, 0, 36892.521, format06); // Jan 1 2001
  worksheet.write(7, 0, 36892.521, format07); // 1 January 2001
  worksheet.write(8, 0, 36892.521, format08); // 01/01/2001 12:30 AM
  worksheet.write(9, 0, 1.87, format09);      // 1 dollar and .87 cents

  // Show limited conditional number formats.
  const xwpp::format_t* format10 = workbook.format_builder().num_format("[Green]General;[Red]-General;General").build();

  worksheet.write(10, 0, 123, format10); // > 0 Green
  worksheet.write(11, 0, -45, format10); // < 0 Red
  worksheet.write(12, 0, 0, format10);   // = 0 Default color

  // Format a Zip code.
  const xwpp::format_t* format11 = workbook.format_builder().num_format("00000").build();

  worksheet.write(13, 0, 1209, format11); // 01209

  workbook.save("format_num_format.xlsx");
}
