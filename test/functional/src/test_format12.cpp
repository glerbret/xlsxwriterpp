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

  const xwpp::format_t* top_left_bottom = workbook.format_builder()
                                            .bottom(xwpp::format_borders_t::THIN)
                                            .left(xwpp::format_borders_t::THIN)
                                            .top(xwpp::format_borders_t::THIN)
                                            .build();

  const xwpp::format_t* top_bottom =
    workbook.format_builder().bottom(xwpp::format_borders_t::THIN).top(xwpp::format_borders_t::THIN).build();

  const xwpp::format_t* top_left =
    workbook.format_builder().left(xwpp::format_borders_t::THIN).top(xwpp::format_borders_t::THIN).build();

  // cppcheck-suppress unreadVariable
  [[maybe_unused]] const xwpp::format_t* unused = workbook.format_builder().left(xwpp::format_borders_t::THIN).build();

  worksheet.write("B2", "test", top_left_bottom);
  worksheet.write("D2", "test", top_left);
  worksheet.write("F2", "test", top_bottom);

  workbook.save("test_format12.xlsx");
}
