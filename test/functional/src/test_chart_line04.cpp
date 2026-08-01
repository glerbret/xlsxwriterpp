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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::LINE);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(47670016, 47671552);

  uint8_t data[5][3] = {
      {5, 10, 15},
      {2, 4,  6 },
      {3, 6,  9 },
      {4, 8,  12},
      {3, 6,  9 }
  };

  for(xwpp::row_num_t row = 0; row < 5; row++)
  {
    for(xwpp::col_num_t col = 0; col < 3; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  xwpp::chart_series_t& series1 = chart.add_series("", "=Sheet1!$A$1:$A$5");
  chart.add_series("", "=Sheet1!$B$1:$B$5");
  xwpp::chart_series_t& series3 = chart.add_series("", "=Sheet1!$C$1:$C$5");

  chart_series_set_smooth(series1, true);
  chart_series_set_smooth(series3, true);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_line04.xlsx");
}
