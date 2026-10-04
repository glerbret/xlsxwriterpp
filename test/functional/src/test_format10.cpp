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

  const xwpp::format_t* border1 = workbook.format_builder().bg_color(xwpp::color_t::red()).build();
  const xwpp::format_t* border2 =
    workbook.format_builder().bg_color(xwpp::color_t::yellow()).pattern(xwpp::format_patterns_t::DARK_VERTICAL).build();
  const xwpp::format_t* border3 = workbook.format_builder()
                                    .bg_color(xwpp::color_t::yellow())
                                    .fg_color(xwpp::color_t::red())
                                    .pattern(xwpp::format_patterns_t::GRAY_0625)
                                    .build();

  worksheet.write_blank(1, 1, border1);
  worksheet.write_blank(3, 1, border2);
  worksheet.write_blank(5, 1, border3);

  workbook.save("test_format10.xlsx");
}
