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
  chart.set_axis_ids(93550464, 93548544);

  worksheet.write_number(0, 0, 10000000);
  worksheet.write_number(1, 0, 20000000);
  worksheet.write_number(2, 0, 30000000);
  worksheet.write_number(3, 0, 20000000);
  worksheet.write_number(4, 0, 10000000);

  worksheet.write_number(0, 1, 10000000);
  worksheet.write_number(1, 1, 20000000);
  worksheet.write_number(2, 1, 30000000);
  worksheet.write_number(3, 1, 20000000);
  worksheet.write_number(4, 1, 10000000);

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");

  chart_axis_set_display_units(chart.y_axis_, xwpp::chart_axis_display_unit_t::HUNDREDS);
  chart_axis_set_display_units_visible(chart.y_axis_, false);

  chart_axis_set_display_units(chart.x_axis_, xwpp::chart_axis_display_unit_t::THOUSANDS);
  chart_axis_set_display_units_visible(chart.x_axis_, false);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_display_units12.xlsx");
}
