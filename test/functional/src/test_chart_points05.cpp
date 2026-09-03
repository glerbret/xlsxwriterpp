/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <array>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::LINE);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(45471616, 46804992);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
  };

  for(xwpp::row_num_t row_num{0}; const auto& row: data)
  {
    for(xwpp::col_num_t col_num{0}; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  xwpp::chart_series_t& series1 = chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");
  xwpp::chart_series_t& series2 = chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$C$1:$C$5");

  const xwpp::chart_fill_t red_fill{.color_ = xwpp::color_t::RED};
  const std::vector<xwpp::chart_point_t> points{{.fill_ = red_fill}};

  series_set_points(series1, points);

  chart.series_set_marker_type(series1, xwpp::chart_marker_type_t::AUTOMATIC);
  chart.series_set_marker_type(series2, xwpp::chart_marker_type_t::AUTOMATIC);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_points05.xlsx");
}
