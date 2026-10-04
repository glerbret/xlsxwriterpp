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

  const xwpp::format_t* border1 =
    workbook.format_builder().border(xwpp::format_borders_t::HAIR).border_color(xwpp::color_t::red()).build();
  const xwpp::format_t* border2 = workbook.format_builder()
                                    .diag_type(xwpp::format_diagonal_types_t::BORDER_UP)
                                    .diag_color(xwpp::color_t::red())
                                    .build();
  const xwpp::format_t* border3 = workbook.format_builder()
                                    .diag_type(xwpp::format_diagonal_types_t::BORDER_DOWN)
                                    .diag_color(xwpp::color_t::red())
                                    .build();
  const xwpp::format_t* border4 = workbook.format_builder()
                                    .diag_type(xwpp::format_diagonal_types_t::BORDER_UP_DOWN)
                                    .diag_color(xwpp::color_t::red())
                                    .build();

  worksheet.write_blank(1, 1, border1);
  worksheet.write_blank(3, 1, border2);
  worksheet.write_blank(5, 1, border3);
  worksheet.write_blank(7, 1, border4);

  workbook.save("test_format09.xlsx");
}
