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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::SCATTER);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(46210048, 46208512);

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

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$C$1:$C$5");

  xwpp::chart_line_t line1 = {.dash_type_ = xwpp::chart_line_dash_type_t::DASH_ROUND_DOT};
  xwpp::chart_line_t line2 = {.dash_type_ = xwpp::chart_line_dash_type_t::DASH_SQUARE_DOT};

  xwpp::chart_fill_t fill1 = {.color_ = xwpp::color_t(0x9999FF)};
  xwpp::chart_fill_t fill2 = {.color_ = xwpp::color_t(0xFFC000)};

  chart.chartarea_set_line(line1);
  chart.chartarea_set_fill(fill1);

  chart.plotarea_set_line(line2);
  chart.plotarea_set_fill(fill2);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_chartarea03.xlsx");
}
