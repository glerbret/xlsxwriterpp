/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();
  xwpp::chart_t& chart          = workbook.add_chart(xwpp::chart_type_t::BAR);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(93218304, 93219840);

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
      worksheet2.write_number(row, col, data[row][col]);
    }
  }

  worksheet1.write_string(CELL("A1"), "Foo");

  chart.add_series("Sheet2!$A$1:$A$5", "Sheet2!$B$1:$B$5");
  chart.add_series("Sheet2!$A$1:$A$5", "Sheet2!$C$1:$C$5");

  worksheet2.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_bar02.xlsx");
}
