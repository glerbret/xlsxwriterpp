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
  chart.set_axis_ids(86421504, 86423040);

  uint8_t data[3][8] = {
      {2, 2, 2, 2, 2, 2, 2, 2},
      {2, 2, 2, 2, 2, 2, 2, 2},
      {2, 2, 2, 2, 2, 2, 2, 2}
  };

  for(xwpp::row_num_t row = 0; row < 3; row++)
  {
    for(xwpp::col_num_t col = 0; col < 8; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  xwpp::chart_series_t& series1 = chart.add_series("", "=Sheet1!$A$1:$A$3");
  xwpp::chart_series_t& series2 = chart.add_series("", "=Sheet1!$B$1:$B$3");
  xwpp::chart_series_t& series3 = chart.add_series("", "=Sheet1!$C$1:$C$3");
  xwpp::chart_series_t& series4 = chart.add_series("", "=Sheet1!$D$1:$D$3");
  xwpp::chart_series_t& series5 = chart.add_series("", "=Sheet1!$E$1:$E$3");
  xwpp::chart_series_t& series6 = chart.add_series("", "=Sheet1!$F$1:$F$3");
  xwpp::chart_series_t& series7 = chart.add_series("", "=Sheet1!$G$1:$G$3");
  xwpp::chart_series_t& series8 = chart.add_series("", "=Sheet1!$H$1:$H$3");

  xwpp::chart_pattern_t pattern1 = {.fg_color_ = xwpp::color_t(0xC00000),
                                    .bg_color_ = xwpp::color_t(0xFFFFFF),
                                    .type_     = xwpp::chart_pattern_type_t::PERCENT_5};
  xwpp::chart_pattern_t pattern2 = {.fg_color_ = xwpp::color_t(0xFF0000),
                                    .type_     = xwpp::chart_pattern_type_t::PERCENT_50};
  xwpp::chart_pattern_t pattern3 = {.fg_color_ = xwpp::color_t(0xFFC000),
                                    .type_     = xwpp::chart_pattern_type_t::LIGHT_DOWNWARD_DIAGONAL};
  xwpp::chart_pattern_t pattern4 = {.fg_color_ = xwpp::color_t(0xFFFF00),
                                    .type_     = xwpp::chart_pattern_type_t::LIGHT_VERTICAL};
  xwpp::chart_pattern_t pattern5 = {.fg_color_ = xwpp::color_t(0x92D050),
                                    .type_     = xwpp::chart_pattern_type_t::DASHED_DOWNWARD_DIAGONAL};
  xwpp::chart_pattern_t pattern6 = {.fg_color_ = xwpp::color_t(0x00B050), .type_ = xwpp::chart_pattern_type_t::ZIGZAG};
  xwpp::chart_pattern_t pattern7 = {.fg_color_ = xwpp::color_t(0x00B0F0), .type_ = xwpp::chart_pattern_type_t::DIVOT};
  xwpp::chart_pattern_t pattern8 = {.fg_color_ = xwpp::color_t(0x0070C0),
                                    .type_     = xwpp::chart_pattern_type_t::SMALL_GRID};

  chart_series_set_pattern(series1, pattern1);
  chart_series_set_pattern(series2, pattern2);
  chart_series_set_pattern(series3, pattern3);
  chart_series_set_pattern(series4, pattern4);
  chart_series_set_pattern(series5, pattern5);
  chart_series_set_pattern(series6, pattern6);
  chart_series_set_pattern(series7, pattern7);
  chart_series_set_pattern(series8, pattern8);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_pattern02.xlsx");
}
