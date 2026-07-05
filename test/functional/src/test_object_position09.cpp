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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::LINE);

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  xwpp::format_t* italic = workbook.add_format();
  italic->set_italic();

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(60910208, 69231360);

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
      worksheet.write_number(row + 1, col, data[row][col]);
    }
  }

  worksheet.write_string(CELL("A1"), "Foo", bold);
  worksheet.write_string(CELL("B1"), "Bar", italic);

  xwpp::row_col_options_t row_hidden = {.hidden_ = true};
  worksheet.set_row(12, xwpp::DEF_ROW_HEIGHT, nullptr, row_hidden);

  xwpp::row_col_options_t col_hidden = {.hidden_ = true};
  worksheet.set_column(COLS("F:F"), 9, nullptr, col_hidden);

  chart.add_series("", "=Sheet1!$A$2:$A$6");
  chart.add_series("", "=Sheet1!$B$2:$B$6");
  chart.add_series("", "=Sheet1!$C$2:$C$6");

  xwpp::chart_options_t chart_options = {.object_position_ = xwpp::object_position_t::MOVE_AND_SIZE_AFTER};
  worksheet.insert_chart(CELL("E9"), &chart, chart_options);

  workbook.save("test_object_position09.xlsx");
}
