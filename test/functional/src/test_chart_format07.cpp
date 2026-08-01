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
  chart.set_axis_ids(46163840, 46175360);

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

  xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$C$1:$C$5");

  xwpp::chart_line_t line = {.color_ = xwpp::color_t::YELLOW};
  xwpp::chart_fill_t fill = {.color_ = xwpp::color_t::RED};

  chart.series_set_marker_type(series, xwpp::chart_marker_type_t::SQUARE);
  chart_series_set_marker_size(series, 5);

  chart_series_set_marker_line(series, line);
  chart_series_set_marker_fill(series, fill);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_format07.xlsx");
}
