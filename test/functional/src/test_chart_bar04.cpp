/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <array>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();
  xwpp::chart_t& chart1         = workbook.add_chart(xwpp::chart_type_t::BAR);
  xwpp::chart_t& chart2         = workbook.add_chart(xwpp::chart_type_t::BAR);

  // For testing, copy the randomly generated axis ids in the target file.
  chart1.set_axis_ids(64446848, 64448384);
  chart2.set_axis_ids(85389696, 85391232);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
  };

  for(xwpp::row_num_t row_num = 0; const auto& row: data)
  {
    for(xwpp::col_num_t col_num = 0; const auto value: row)
    {
      worksheet1.write_number(row_num, col_num, value);
      worksheet2.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  chart1.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");
  chart1.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$C$1:$C$5");

  worksheet1.insert_chart(CELL("E9"), &chart1);

  chart2.add_series("=Sheet2!$A$1:$A$5", "=Sheet2!$B$1:$B$5");
  chart2.add_series("=Sheet2!$A$1:$A$5", "=Sheet2!$C$1:$C$5");

  worksheet2.insert_chart(CELL("E9"), &chart2);

  workbook.save("test_chart_bar04.xlsx");
}
