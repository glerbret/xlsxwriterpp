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

  for(xwpp::row_num_t row_num = 0; const auto& row: data)
  {
    for(xwpp::col_num_t col_num = 0; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  xwpp::chart_series_t& series1 = chart.add_series("", "=Sheet1!$A$1:$A$3");
  xwpp::chart_series_t& series2 = chart.add_series("", "=Sheet1!$B$1:$B$3");
  xwpp::chart_series_t& series3 = chart.add_series("", "=Sheet1!$C$1:$C$3");
  xwpp::chart_series_t& series4 = chart.add_series("", "=Sheet1!$D$1:$D$3");
  xwpp::chart_series_t& series5 = chart.add_series("", "=Sheet1!$E$1:$E$3");
  xwpp::chart_series_t& series6 = chart.add_series("", "=Sheet1!$F$1:$F$3");
  xwpp::chart_series_t& series7 = chart.add_series("", "=Sheet1!$G$1:$G$3");
  xwpp::chart_series_t& series8 = chart.add_series("", "=Sheet1!$H$1:$H$3");

  const xwpp::chart_pattern_t pattern1{.fg_color_ = xwpp::color_t(0xC00000),
                                       .bg_color_ = xwpp::color_t(0xFFFFFF),
                                       .type_     = xwpp::chart_pattern_type_t::PERCENT_30};
  const xwpp::chart_pattern_t pattern2{.fg_color_ = xwpp::color_t(0xFF0000),
                                       .type_     = xwpp::chart_pattern_type_t::PERCENT_80};
  const xwpp::chart_pattern_t pattern3{.fg_color_ = xwpp::color_t(0xFFC000),
                                       .type_     = xwpp::chart_pattern_type_t::WIDE_DOWNWARD_DIAGONAL};
  const xwpp::chart_pattern_t pattern4{.fg_color_ = xwpp::color_t(0xFFFF00),
                                       .type_     = xwpp::chart_pattern_type_t::DARK_VERTICAL};
  const xwpp::chart_pattern_t pattern5{.fg_color_ = xwpp::color_t(0x92D050),
                                       .type_     = xwpp::chart_pattern_type_t::SMALL_CONFETTI};
  const xwpp::chart_pattern_t pattern6{.fg_color_ = xwpp::color_t(0x00B050),
                                       .type_     = xwpp::chart_pattern_type_t::WEAVE};
  const xwpp::chart_pattern_t pattern7{.fg_color_ = xwpp::color_t(0x00B0F0),
                                       .type_     = xwpp::chart_pattern_type_t::TRELLIS};
  const xwpp::chart_pattern_t pattern8{.fg_color_ = xwpp::color_t(0x0070C0),
                                       .type_     = xwpp::chart_pattern_type_t::OUTLINED_DIAMOND};

  chart_series_set_pattern(series1, pattern1);
  chart_series_set_pattern(series2, pattern2);
  chart_series_set_pattern(series3, pattern3);
  chart_series_set_pattern(series4, pattern4);
  chart_series_set_pattern(series5, pattern5);
  chart_series_set_pattern(series6, pattern6);
  chart_series_set_pattern(series7, pattern7);
  chart_series_set_pattern(series8, pattern8);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_pattern06.xlsx");
}
