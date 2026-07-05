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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::PIE);

  uint8_t data[3][2] = {
      {2, 60},
      {4, 30},
      {6, 10},
  };

  for(size_t row = 0; row < 3; row++)
  {
    for(size_t col = 0; col < 2; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  chart.add_series("=Sheet1!$A$1:$A$3", "=Sheet1!$B$1:$B$3");

  chart.title_set_name("Title");

  // Axis formatting should be igmored.
  xwpp::chart_axis_set_name(chart.x_axis_, "XXX");
  xwpp::chart_axis_set_name(chart.y_axis_, "YYY");

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_axis06.xlsx");
}
