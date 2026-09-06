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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::LINE);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(49019136, 49222016);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 5}, {4, 8, 12}, {5, 10, 15}}
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

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$C$1:$C$5");

  const xwpp::chart_line_t up_line{.color_ = xwpp::color_t::YELLOW};
  const xwpp::chart_fill_t up_fill{.color_ = xwpp::color_t::RED};
  const xwpp::chart_line_t down_line{.color_     = xwpp::color_t(0x00B0F0),
                                     .dash_type_ = xwpp::chart_line_dash_type_t::SQUARE_DOT};
  const xwpp::chart_fill_t down_fill{.color_ = xwpp::color_t(0x00B050)};

  chart.set_up_down_bars_format(up_line, up_fill, down_line, down_fill);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_up_down_bars02.xlsx");
}
