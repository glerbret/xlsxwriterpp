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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::SCATTER);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(45740416, 45705856);

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

  xwpp::chart_series_t& series1 = chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");
  xwpp::chart_series_t& series2 = chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$C$1:$C$5");

  xwpp::chart_series_set_labels(series1);
  xwpp::chart_series_set_labels(series2);
  xwpp::chart_series_set_labels_position(series1, xwpp::chart_label_position_t::LEFT);
  xwpp::chart_series_set_labels_position(series2, xwpp::chart_label_position_t::CENTER);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_data_labels10.xlsx");
}
