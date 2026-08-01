/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::DOUGHNUT);

  uint8_t data[3][2] = {
      {2, 60},
      {4, 30},
      {6, 10},
  };

  for(xwpp::row_num_t row = 0; row < 3; row++)
  {
    for(xwpp::col_num_t col = 0; col < 2; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  chart.add_series("=Sheet1!$A$1:$A$3", "=Sheet1!$B$1:$B$3");

  xwpp::chart_line_t line1 = {.color_     = xwpp::color_t(0xFFFF00),
                              .dash_type_ = xwpp::chart_line_dash_type_t::DASH_LONG_DASH};
  xwpp::chart_line_t line2 = {.dash_type_ = xwpp::chart_line_dash_type_t::DASH_SQUARE_DOT};

  xwpp::chart_fill_t fill1 = {.color_ = xwpp::color_t(0x92D050)};
  xwpp::chart_fill_t fill2 = {.color_ = xwpp::color_t(0xFF0000)};

  chart.chartarea_set_line(line1);
  chart.chartarea_set_fill(fill1);

  chart.plotarea_set_line(line2);
  chart.plotarea_set_fill(fill2);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_chartarea06.xlsx");
}
