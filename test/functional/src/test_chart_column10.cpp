/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <tuple>
#include <vector>

using namespace std::string_literals;

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(45686144, 45722240);

  const std::vector<std::tuple<std::string, int>> data{
    {"A"s, 1},
    {"B"s, 2},
    {"C"s, 3},
    {"D"s, 2},
    {"E"s, 1},
  };

  for(xwpp::row_num_t row_num = 0; const auto& value: data)
  {
    worksheet.write_string(row_num, 0, std::get<std::string>(value));
    worksheet.write_number(row_num, 1, std::get<int>(value));
    row_num++;
  }

  chart.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_column10.xlsx");
}
