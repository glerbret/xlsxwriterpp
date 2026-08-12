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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::SCATTER);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(41671680, 41669376);

  worksheet.write_number(0, 0, 1);
  worksheet.write_number(1, 0, 2);
  worksheet.write_string(2, 0, "Foo");
  worksheet.write_number(3, 0, 4);
  worksheet.write_number(4, 0, 5);

  worksheet.write_number(0, 1, 2);
  worksheet.write_number(1, 1, 4);
  worksheet.write_string(2, 1, "Bar");
  worksheet.write_number(3, 1, 8);
  worksheet.write_number(4, 1, 10);

  worksheet.write_number(0, 2, 3);
  worksheet.write_number(1, 2, 6);
  worksheet.write_string(2, 2, "Baz");
  worksheet.write_number(3, 2, 12);
  worksheet.write_number(4, 2, 15);

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$C$1:$C$5");
  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_str02.xlsx");
}
