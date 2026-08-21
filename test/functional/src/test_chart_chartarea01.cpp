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
  chart.set_axis_ids(82933248, 82952960);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
  };

  for(xwpp::row_num_t row_num = 0; const auto& row: data)
  {
    for(xwpp::col_num_t col_num = 0; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  chart.add_series("", "=Sheet1!$A$1:$A$5");
  chart.add_series("", "=Sheet1!$B$1:$B$5");
  chart.add_series("", "=Sheet1!$C$1:$C$5");

  const xwpp::chart_line_t line1 = {.none_ = true};
  const xwpp::chart_line_t line2 = {
    .color_ = xwpp::color_t::YELLOW, .width_ = 1, .dash_type_ = xwpp::chart_line_dash_type_t::DASH};

  const xwpp::chart_fill_t fill1 = {.color_ = xwpp::color_t::RED};
  const xwpp::chart_fill_t fill2 = {.color_ = xwpp::color_t(0x92D050)};

  chart.chartarea_set_line(line1);
  chart.chartarea_set_fill(fill1);

  chart.plotarea_set_line(line2);
  chart.plotarea_set_fill(fill2);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_chartarea01.xlsx");
}
