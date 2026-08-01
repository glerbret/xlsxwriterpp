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

  chart.add_series("", "=Sheet1!$A$1:$A$3");
  chart.add_series("", "=Sheet1!$B$1:$B$3");
  chart.add_series("", "=Sheet1!$C$1:$C$3");
  chart.add_series("", "=Sheet1!$D$1:$D$3");
  chart.add_series("", "=Sheet1!$E$1:$E$3");
  chart.add_series("", "=Sheet1!$F$1:$F$3");
  chart.add_series("", "=Sheet1!$G$1:$G$3");
  chart.add_series("", "=Sheet1!$H$1:$H$3");

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_pattern01.xlsx");
}
