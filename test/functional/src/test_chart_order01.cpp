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
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();
  xwpp::chart_t& chart1         = workbook.add_chart(xwpp::chart_type_t::COLUMN);
  xwpp::chart_t& chart2         = workbook.add_chart(xwpp::chart_type_t::BAR);
  xwpp::chart_t& chart3         = workbook.add_chart(xwpp::chart_type_t::LINE);
  xwpp::chart_t& chart4         = workbook.add_chart(xwpp::chart_type_t::PIE);

  // For testing, copy the randomly generated axis ids in the target file.
  chart1.set_axis_ids(54976896, 54978432);
  chart2.set_axis_ids(54310784, 54312320);
  chart3.set_axis_ids(69816704, 69818240);
  chart4.set_axis_ids(69816704, 69818240);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
  };

  for(xwpp::row_num_t row_num = 0; const auto& row: data)
  {
    for(xwpp::col_num_t col_num = 0; const auto value: row)
    {
      worksheet1.write_number(row_num, col_num, value);
      worksheet2.write_number(row_num, col_num, value);
      worksheet3.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  chart1.add_series("", "=Sheet1!$A$1:$A$5");
  chart2.add_series("", "=Sheet2!$A$1:$A$5");
  chart3.add_series("", "=Sheet3!$A$1:$A$5");
  chart4.add_series("", "=Sheet1!$B$1:$B$5");

  worksheet1.insert_chart(CELL("E9"), &chart1);
  worksheet2.insert_chart(CELL("E9"), &chart2);
  worksheet3.insert_chart(CELL("E9"), &chart3);
  worksheet1.insert_chart(CELL("E24"), &chart4);

  workbook.save("test_chart_order01.xlsx");
}
