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

  xwpp::format_t* bold_italic = workbook.add_format();
  bold_italic->set_bold();
  bold_italic->set_italic();

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(68311296, 69198208);

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
      worksheet.write_number(row + 1, col, data[row][col], nullptr);
    }
  }

  worksheet.write_string(CELL("A1"), "Foo", italic);
  worksheet.write_string(CELL("B1"), "Bar", bold);

  worksheet.set_row(12, xwpp::DEF_ROW_HEIGHT, italic);
  worksheet.set_column(COLS("F:F"), xwpp::DEF_COL_WIDTH, bold);

  worksheet.write_blank(CELL("F13"), bold_italic);

  chart.add_series("", "=Sheet1!$A$2:$A$6");
  chart.add_series("", "=Sheet1!$B$2:$B$6");
  chart.add_series("", "=Sheet1!$C$2:$C$6");

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_set_column05.xlsx");
}
