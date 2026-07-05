/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(60474496, 78612736);

  worksheet.write_string(CELL("A1"), "1.1_1");
  worksheet.write_string(CELL("B1"), "2.2_2");
  worksheet.write_number(CELL("A2"), 1);
  worksheet.write_number(CELL("B2"), 2);

  chart.add_series("=Sheet1!$A$1:$B$1", "=Sheet1!$A$2:$B$2");

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_column13.xlsx");
}
