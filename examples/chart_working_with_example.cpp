/*
 * An example of a simple Excel chart using the Xlsxwriter++ library. This
 * example is used in the "Working with Charts" section of the docs.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Write some data for the chart.
  worksheet.write_number(0, 0, 10);
  worksheet.write_number(1, 0, 40);
  worksheet.write_number(2, 0, 50);
  worksheet.write_number(3, 0, 20);
  worksheet.write_number(4, 0, 10);
  worksheet.write_number(5, 0, 50);

  // Create a chart object.
  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Configure the chart.
  const xwpp::chart_series_t& series = chart.add_series("", "Sheet1!$A$1:$A$6");
  (void)series; // Do something with series in the real examples.

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("C1"), &chart);

  workbook.save("chart_working_with_example.xlsx");
}
