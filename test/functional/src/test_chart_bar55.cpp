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
  chart.set_axis_ids(64264064, 64447232);

  uint8_t data[5][3] = {
      {1, 2,  3 },
      {2, 4,  6 },
      {3, 6,  9 },
      {4, 8,  12},
      {5, 10, 15}
  };

  for(xwpp::row_num_t row = 0; row < 5; row++)
  {
    for(xwpp::col_num_t col = 0; col < 3; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  xwpp::chart_series_t& series1 = chart.add_series("", "=Sheet1!$A$1:$A$5");
  xwpp::chart_series_t& series2 = chart.add_series("", "=Sheet1!$B$1:$B$5");
  xwpp::chart_series_t& series3 = chart.add_series("", "=Sheet1!$C$1:$C$5");

  // Add cache data for testing.
  xwpp::chart_add_data_cache(series1.values_, data[0], 5, 3, 0);
  xwpp::chart_add_data_cache(series2.values_, data[0], 5, 3, 1);
  xwpp::chart_add_data_cache(series3.values_, data[0], 5, 3, 2);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_bar55.xlsx");
}
