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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(45472384, 49016832);

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

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$C$1:$C$5");
  chart.series_set_error_bars(series1.y_error_bars_, xwpp::chart_error_bar_type_t::STD_ERROR, 0);
  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_errorbars06.xlsx");
}
