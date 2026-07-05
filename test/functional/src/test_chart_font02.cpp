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
  chart.set_axis_ids(43945344, 45705856);

  uint8_t data[5][3] = {
      {1, 2,  3 },
      {2, 4,  6 },
      {3, 6,  9 },
      {4, 8,  12},
      {5, 10, 15}
  };

  for(size_t row = 0; row < 5; row++)
  {
    for(size_t col = 0; col < 3; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  chart.add_series("", "=Sheet1!$A$1:$A$5");
  chart.add_series("", "=Sheet1!$B$1:$B$5");
  chart.add_series("", "=Sheet1!$C$1:$C$5");

  chart.title_set_name("Title");

  xwpp::chart_font_t font1 = {.italic_ = true};
  xwpp::chart_font_t font2 = {.bold_ = true, .italic_ = true};
  xwpp::chart_font_t font3 = {.size_ = 11, .bold_ = true, .italic_ = true};
  xwpp::chart_font_t font4 = {.size_ = 9, .italic_ = true};

  chart_axis_set_name(chart.x_axis_, "XXX");
  chart_axis_set_name_font(chart.x_axis_, font1);

  chart_axis_set_name(chart.y_axis_, "YYY");
  chart_axis_set_name_font(chart.y_axis_, font2);

  chart_axis_set_num_font(chart.x_axis_, font3);
  chart_axis_set_num_font(chart.y_axis_, font4);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_font02.xlsx");
}
