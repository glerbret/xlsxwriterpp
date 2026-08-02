/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  /* xwpp::worksheet_t& worksheet1 =*/workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet3 =*/workbook.add_worksheet();
  xwpp::chartsheet_t& chartsheet1 = workbook.add_chartsheet();
  xwpp::chart_t& chart1           = workbook.add_chart(xwpp::chart_type_t::BAR);
  xwpp::chart_t& chart2           = workbook.add_chart(xwpp::chart_type_t::BAR);
  xwpp::chart_t& chart3           = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart1.set_axis_ids(40294272, 40295808);
  chart2.set_axis_ids(40261504, 65749760);
  chart3.set_axis_ids(65465728, 66388352);

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
      worksheet2.write_number(row, col, data[row][col]);
    }
  }

  worksheet2.write_url(CELL("A6"), "http://www.perl.com/");

  chart3.add_series("", "=Sheet2!$A$1:$A$5");
  chart3.add_series("", "=Sheet2!$B$1:$B$5");
  chart3.add_series("", "=Sheet2!$C$1:$C$5");

  chart1.add_series("", "=Sheet2!$A$1:$A$5");
  chart1.add_series("", "=Sheet2!$B$1:$B$5");
  chart1.add_series("", "=Sheet2!$C$1:$C$5");

  chart2.add_series("", "=Sheet2!$A$1:$A$5");

  worksheet2.insert_chart(CELL("E9"), &chart1);
  worksheet2.insert_chart(CELL("F25"), &chart2);

  chartsheet1.set_chart(&chart3);

  workbook.save("test_chart_bar14.xlsx");
}
