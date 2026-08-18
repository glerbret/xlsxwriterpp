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
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::LINE);

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  xwpp::format_t* italic = workbook.add_format();
  italic->set_italic();

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(60888960, 79670656);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
  };

  for(xwpp::row_num_t row_num = 1; const auto& row: data)
  {
    for(xwpp::col_num_t col_num = 0; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  worksheet.write_string(CELL("A1"), "Foo", bold);
  worksheet.write_string(CELL("B1"), "Bar", italic);

  const xwpp::row_col_options_t row_hidden{.hidden_ = true};
  worksheet.set_row(12, xwpp::DEF_ROW_HEIGHT, nullptr, row_hidden);

  const xwpp::row_col_options_t col_hidden{.hidden_ = true};
  worksheet.set_column(COLS("F:F"), xwpp::DEF_COL_WIDTH, nullptr, col_hidden);

  chart.add_series("", "=Sheet1!$A$2:$A$6");
  chart.add_series("", "=Sheet1!$B$2:$B$6");
  chart.add_series("", "=Sheet1!$C$2:$C$6");

  const xwpp::chart_options_t chart_options{.object_position_ = xwpp::object_position_t::MOVE_DONT_SIZE};
  worksheet.insert_chart(CELL("E9"), &chart, chart_options);

  workbook.save("test_object_position08.xlsx");
}
