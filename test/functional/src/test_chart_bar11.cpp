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
  xwpp::chart_t& chart1        = workbook.add_chart(xwpp::chart_type_t::BAR);
  xwpp::chart_t& chart2        = workbook.add_chart(xwpp::chart_type_t::BAR);
  xwpp::chart_t& chart3        = workbook.add_chart(xwpp::chart_type_t::BAR);

  // For testing, copy the randomly generated axis ids in the target file.
  chart1.set_axis_ids(40274944, 40294272);
  chart2.set_axis_ids(62355328, 62356864);
  chart3.set_axis_ids(79538816, 65422464);

  workbook.unset_default_url_format();

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

  worksheet.write_url(CELL("A7"), "http://www.perl.com/");
  worksheet.write_url(CELL("A8"), "http://www.perl.org/");
  worksheet.write_url(CELL("A9"), "http://www.perl.net/");

  chart1.add_series("", "=Sheet1!$A$1:$A$5");
  chart1.add_series("", "=Sheet1!$B$1:$B$5");
  chart1.add_series("", "=Sheet1!$C$1:$C$5");

  chart2.add_series("", "=Sheet1!$A$1:$A$5");
  chart2.add_series("", "=Sheet1!$B$1:$B$5");

  chart3.add_series("", "=Sheet1!$A$1:$A$5");

  worksheet.insert_chart(CELL("E9"), &chart1);
  worksheet.insert_chart(CELL("D25"), &chart2);
  worksheet.insert_chart(CELL("L32"), &chart3);

  workbook.save("test_chart_bar11.xlsx");
}
