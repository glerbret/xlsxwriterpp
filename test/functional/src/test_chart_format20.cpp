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
  xwpp::chart_t& chart1        = workbook.add_chart(xwpp::chart_type_t::LINE);
  xwpp::chart_t& chart2        = workbook.add_chart(xwpp::chart_type_t::LINE);

  // For testing, copy the randomly generated axis ids in the target file.
  chart1.set_axis_ids(80553856, 80555392);
  chart2.set_axis_ids(84583936, 84585856);

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

  xwpp::chart_series_t& series1 = chart1.add_series("", "=Sheet1!$B$1:$B$5");
  chart1.add_series("", "=Sheet1!$C$1:$C$5");

  xwpp::chart_series_t& series2 = chart2.add_series("", "=Sheet1!$B$1:$B$5");
  chart2.add_series("", "=Sheet1!$C$1:$C$5");

  xwpp::chart_line_t line = {.color_ = xwpp::color_t::RED, .dash_type_ = xwpp::chart_line_dash_type_t::DASH_DASH};

  series_set_trendline(series1, xwpp::chart_trendline_type_t::LINEAR, 0);
  series_set_trendline_line(series1, line);
  series_set_trendline(series2, xwpp::chart_trendline_type_t::LINEAR, 0);
  series_set_trendline_line(series2, line);

  worksheet.insert_chart(CELL("E9"), &chart1);
  worksheet.insert_chart(CELL("E25"), &chart2);

  workbook.save("test_chart_format20.xlsx");
}
