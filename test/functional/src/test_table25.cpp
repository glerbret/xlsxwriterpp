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

  worksheet.set_column(COLS("C:F"), 10.288, nullptr);

  const xwpp::table_options_t options{.style_type_ = xwpp::table_style_type_t::LIGHT, .style_type_number_ = 0};

  worksheet.add_table(RANGE("C3:F13"), options);

  workbook.save("test_table25.xlsx");
}
