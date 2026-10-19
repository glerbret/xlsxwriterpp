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
  chart.set_axis_ids(110902272, 110756608);

  const std::array<std::array<uint8_t, 8>, 3> data{
    {{2, 2, 2, 2, 2, 2, 2, 2}, {2, 2, 2, 2, 2, 2, 2, 2}, {2, 2, 2, 2, 2, 2, 2, 2}}
  };

  for(xwpp::row_num_t row_num{0}; const auto& row: data)
  {
    for(xwpp::col_num_t col_num{0}; const auto value: row)
    {
      worksheet.write(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }


  const xwpp::chart_pattern_t pattern1{.fg_color = xwpp::color_t{0xC00000},
                                       .bg_color = xwpp::color_t{0xFFFFFF},
                                       .type     = xwpp::chart_pattern_type_t::PERCENT_10};
  chart.add_series("=Sheet1!$A$1:$A$3").set_pattern(pattern1);

  const xwpp::chart_pattern_t pattern2{.fg_color = xwpp::color_t{0xFF0000},
                                       .type     = xwpp::chart_pattern_type_t::PERCENT_60};
  chart.add_series("=Sheet1!$B$1:$B$3").set_pattern(pattern2);

  const xwpp::chart_pattern_t pattern3{.fg_color = xwpp::color_t{0xFFC000},
                                       .type     = xwpp::chart_pattern_type_t::LIGHT_UPWARD_DIAGONAL};
  chart.add_series("=Sheet1!$C$1:$C$3").set_pattern(pattern3);

  const xwpp::chart_pattern_t pattern4{.fg_color = xwpp::color_t{0xFFFF00},
                                       .type     = xwpp::chart_pattern_type_t::LIGHT_HORIZONTAL};
  chart.add_series("=Sheet1!$D$1:$D$3").set_pattern(pattern4);

  const xwpp::chart_pattern_t pattern5{.fg_color = xwpp::color_t{0x92D050},
                                       .type     = xwpp::chart_pattern_type_t::DASHED_UPWARD_DIAGONAL};
  chart.add_series("=Sheet1!$E$1:$E$3").set_pattern(pattern5);

  const xwpp::chart_pattern_t pattern6{.fg_color = xwpp::color_t{0x00B050}, .type = xwpp::chart_pattern_type_t::WAVE};
  chart.add_series("=Sheet1!$F$1:$F$3").set_pattern(pattern6);

  const xwpp::chart_pattern_t pattern7{.fg_color = xwpp::color_t{0x00B0F0},
                                       .type     = xwpp::chart_pattern_type_t::DOTTED_GRID};
  chart.add_series("=Sheet1!$G$1:$G$3").set_pattern(pattern7);

  const xwpp::chart_pattern_t pattern8{.fg_color = xwpp::color_t{0x0070C0},
                                       .type     = xwpp::chart_pattern_type_t::LARGE_GRID};
  chart.add_series("=Sheet1!$H$1:$H$3").set_pattern(pattern8);

  worksheet.insert_chart("E9", &chart);

  workbook.save("test_chart_pattern03.xlsx");
}
