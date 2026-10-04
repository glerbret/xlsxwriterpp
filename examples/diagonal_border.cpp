/*
 * A simple formatting example that demonstrates how to add diagonal
 * cell borders using the Xlsxwriter++ library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add some diagonal border formats.
  const xwpp::format_t* format1 = workbook.format_builder().diag_type(xwpp::format_diagonal_types_t::BORDER_UP).build();
  const xwpp::format_t* format2 =
    workbook.format_builder().diag_type(xwpp::format_diagonal_types_t::BORDER_DOWN).build();
  const xwpp::format_t* format3 =
    workbook.format_builder().diag_type(xwpp::format_diagonal_types_t::BORDER_UP_DOWN).build();
  const xwpp::format_t* format4 = workbook.format_builder()
                                    .diag_type(xwpp::format_diagonal_types_t::BORDER_UP_DOWN)
                                    .diag_border(xwpp::format_borders_t::HAIR)
                                    .diag_color(xwpp::color_t::red())
                                    .build();

  worksheet.write("B3", "Text", format1);
  worksheet.write("B6", "Text", format2);
  worksheet.write("B9", "Text", format3);
  worksheet.write("B12", "Text", format4);

  workbook.save("diagonal_border.xlsx");
}
