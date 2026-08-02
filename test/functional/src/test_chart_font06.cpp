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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::BAR);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(49407488, 53740288);

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

  chart.add_series("", "=Sheet1!$A$1:$A$5");
  chart.add_series("", "=Sheet1!$B$1:$B$5");
  chart.add_series("", "=Sheet1!$C$1:$C$5");

  xwpp::chart_font_t font1 = {.name_ = "Calibri", .color_ = xwpp::color_t::YELLOW, .pitch_family_ = 34, .charset_ = 0};
  xwpp::chart_font_t font2 = {
    .name_ = "Courier New", .color_ = xwpp::color_t(0x92D050), .pitch_family_ = 49, .charset_ = 0};
  xwpp::chart_font_t font3 = {.name_ = "Arial", .color_ = xwpp::color_t(0x00B0F0), .pitch_family_ = 34, .charset_ = 0};
  xwpp::chart_font_t font4 = {.name_ = "Century", .color_ = xwpp::color_t::RED, .pitch_family_ = 18, .charset_ = 0};
  xwpp::chart_font_t font5 = {.bold_ = true, .italic_ = true, .underline_ = true, .color_ = xwpp::color_t(0x7030A0)};

  chart.title_set_name("Title");
  chart.title_set_name_font(font1);

  chart_axis_set_name(chart.x_axis_, "XXX");
  chart_axis_set_name_font(chart.x_axis_, font2);
  chart_axis_set_num_font(chart.x_axis_, font3);

  chart_axis_set_name(chart.y_axis_, "YYY");
  chart_axis_set_name_font(chart.y_axis_, font4);
  chart_axis_set_num_font(chart.y_axis_, font5);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_font06.xlsx");
}
