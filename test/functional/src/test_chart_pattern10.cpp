/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <array>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(143227136, 143245312);

  const std::array<std::array<uint8_t, 8>, 3> data{
    {{2, 2, 2, 2, 2, 2, 2, 2}, {2, 2, 2, 2, 2, 2, 2, 2}, {2, 2, 2, 2, 2, 2, 2, 2}}
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

  chart.add_series("", "=Sheet1!$A$1:$A$3");
  chart.add_series("", "=Sheet1!$B$1:$B$3");
  chart.add_series("", "=Sheet1!$C$1:$C$3");
  chart.add_series("", "=Sheet1!$D$1:$D$3");
  chart.add_series("", "=Sheet1!$E$1:$E$3");
  chart.add_series("", "=Sheet1!$F$1:$F$3");
  chart.add_series("", "=Sheet1!$G$1:$G$3");
  chart.add_series("", "=Sheet1!$H$1:$H$3");

  const xwpp::chart_pattern_t pattern{.fg_color_ = xwpp::color_t::RED,
                                      .bg_color_ = xwpp::color_t::YELLOW,
                                      .type_     = xwpp::chart_pattern_type_t::PERCENT_5};

  chart.plotarea_set_pattern(pattern);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_pattern10.xlsx");
}
