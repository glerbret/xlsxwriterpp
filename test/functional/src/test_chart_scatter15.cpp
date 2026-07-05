/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::SCATTER);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(58843520, 58845440);

  worksheet.write_string(0, 0, "X", NULL);
  worksheet.write_string(0, 1, "Y", NULL);
  worksheet.write_number(1, 0, 1, NULL);
  worksheet.write_number(1, 1, 10, NULL);
  worksheet.write_number(2, 0, 3, NULL);
  worksheet.write_number(2, 1, 30, NULL);

  chart.add_series("=Sheet1!$A$2:$A$3", "=Sheet1!$B$2:$B$3");

  xwpp::chart_font_t font1 = {.italic_ = true, .baseline_ = -1};
  chart_axis_set_name_range(chart.x_axis_, "Sheet1", 0, 0);
  chart_axis_set_name_font(chart.x_axis_, font1);

  chart_axis_set_name_range(chart.y_axis_, "Sheet1", 0, 1);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_scatter15.xlsx");
}
