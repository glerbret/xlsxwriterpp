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

  worksheet.set_column(COLS("C:F"), 10.288);

  xwpp::table_options_t options = {.style_type_ = xwpp::table_style_type_t::MEDIUM, .style_type_number_ = 10};

  worksheet.add_table(RANGE("C3:F13"), options);

  workbook.save("test_table24.xlsx");
}
