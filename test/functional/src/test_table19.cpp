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

  std::vector<xwpp::table_column_t> columns = {{}, {}, {}, {.header_ = " Column4 "}};
  xwpp::table_options_t options             = {.columns_ = columns};

  worksheet.add_table(RANGE("C3:F13"), options);

  workbook.save("test_table19.xlsx");
}
