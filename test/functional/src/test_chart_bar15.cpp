/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1   = workbook.add_worksheet();
  xwpp::chartsheet_t& chartsheet1 = workbook.add_chartsheet();
  xwpp::worksheet_t& worksheet2   = workbook.add_worksheet();
  xwpp::chartsheet_t& chartsheet2 = workbook.add_chartsheet();
  xwpp::chart_t& chart1           = workbook.add_chart(xwpp::chart_type_t::BAR);
  xwpp::chart_t& chart2           = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart1.set_axis_ids(62576896, 62582784);
  chart2.set_axis_ids(65979904, 65981440);

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
      worksheet1.write_number(row, col, data[row][col]);
      worksheet2.write_number(row, col, data[row][col]);
    }
  }

  chart1.add_series("", "=Sheet1!$A$1:$A$5");
  chart1.add_series("", "=Sheet1!$B$1:$B$5");
  chart1.add_series("", "=Sheet1!$C$1:$C$5");

  chart2.add_series("", "=Sheet2!$A$1:$A$5");

  chartsheet1.set_chart(&chart1);
  chartsheet2.set_chart(&chart2);

  workbook.save("test_chart_bar15.xlsx");
}
