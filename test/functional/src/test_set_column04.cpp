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

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  xwpp::format_t* italic = workbook.add_format();
  italic->set_italic();

  xwpp::format_t* bold_italic = workbook.add_format();
  bold_italic->set_bold();
  bold_italic->set_italic();

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

  worksheet.write_string(CELL("A1"), "Foo", italic);
  worksheet.write_string(CELL("B1"), "Bar", bold);

  worksheet.set_row(12, xwpp::DEF_ROW_HEIGHT, italic);
  worksheet.set_column(COLS("F:F"), xwpp::DEF_COL_WIDTH, bold);

  worksheet.write_blank(CELL("F13"), bold_italic);

  workbook.save("test_set_column04.xlsx");
}
