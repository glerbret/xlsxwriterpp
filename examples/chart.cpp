/*
 * An example of a simple Excel chart using the Xlsxwriter++ library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <array>

namespace
{

void write_worksheet_data(xwpp::worksheet_t& worksheet)
{
  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
  };

  for(xwpp::row_num_t row_num{0}; const auto& row: data)
  {
    for(xwpp::col_num_t col_num{0}; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }
}

}

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Write some data for the chart.
  write_worksheet_data(worksheet);

  // Create a chart object.
  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // Configure the chart. In simplest case we just add some value data
  // series. The nullptr categories will default to 1 to 5 like in Excel.
  chart.add_series("", "Sheet1!$A$1:$A$5");
  chart.add_series("", "Sheet1!$B$1:$B$5");
  chart.add_series("", "Sheet1!$C$1:$C$5");

  const xwpp::chart_font_t font{.bold_ = false, .color_ = xwpp::color_t::BLUE};

  chart.title_set_name("Year End Results");
  chart.title_set_name_font(font);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("B7"), &chart);

  workbook.save("chart.xlsx");
}
