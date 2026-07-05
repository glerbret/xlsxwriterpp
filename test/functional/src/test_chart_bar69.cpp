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
  chart.set_axis_ids(66558592, 66569344);

  uint8_t data[5][3] = {
      {1, 2,  3 },
      {2, 4,  6 },
      {3, 6,  9 },
      {4, 8,  12},
      {5, 10, 15}
  };

  for(size_t row = 0; row < 5; row++)
  {
    for(size_t col = 0; col < 3; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  xwpp::chart_series_t& series1 = chart.add_series("", "");
  xwpp::chart_series_t& series2 = chart.add_series("", "");
  xwpp::chart_series_t& series3 = chart.add_series("", "");

  xwpp::chart_series_set_values(series1, "Sheet1", 0, 0, 4, 0);
  xwpp::chart_series_set_values(series2, "Sheet1", 0, 1, 4, 1);
  xwpp::chart_series_set_values(series3, "Sheet1", 0, 2, 4, 2);

  xwpp::chart_axis_set_name_range(chart.x_axis_, "Sheet1", 1, 0);
  xwpp::chart_axis_set_name_range(chart.y_axis_, "Sheet1", 2, 0);
  chart.title_set_name_range("Sheet1", 0, 0);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_bar69.xlsx");
}
