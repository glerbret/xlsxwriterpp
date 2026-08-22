/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::BAR);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(43706240, 43727104);

  worksheet.write_string(0, 1, "Series 1");
  worksheet.write_string(0, 2, "Series 2");
  worksheet.write_string(0, 3, "Series 3");

  worksheet.write_string(1, 0, "Category 1");
  worksheet.write_string(2, 0, "Category 2");
  worksheet.write_string(3, 0, "Category 3");
  worksheet.write_string(4, 0, "Category 4");

  worksheet.write_number(1, 1, 4.3);
  worksheet.write_number(2, 1, 2.5);
  worksheet.write_number(3, 1, 3.5);
  worksheet.write_number(4, 1, 4.5);

  worksheet.write_number(1, 2, 2.4);
  worksheet.write_number(2, 2, 4.5);
  worksheet.write_number(3, 2, 1.8);
  worksheet.write_number(4, 2, 2.8);

  worksheet.write_number(1, 3, 2);
  worksheet.write_number(2, 3, 2);
  worksheet.write_number(3, 3, 3);
  worksheet.write_number(4, 3, 5);

  worksheet.set_column(COLS("A:D"), 11);

  chart.add_series("=Sheet1!$A$2:$A$5", "=Sheet1!$B$2:$B$5");
  chart.add_series("=Sheet1!$A$2:$A$5", "=Sheet1!$C$2:$C$5");
  chart.add_series("=Sheet1!$A$2:$A$5", "=Sheet1!$D$2:$D$5");

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_bar22.xlsx");
}
