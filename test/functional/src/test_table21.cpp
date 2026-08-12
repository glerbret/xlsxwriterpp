/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.set_column(COLS("C:D"), 10.288, nullptr);

  worksheet.write_string(CELL("A1"), "Column", nullptr);

  const std::vector<xwpp::table_column_t> columns{{.header_ = "Column"}};

  const xwpp::table_options_t options{.columns_ = columns};

  worksheet.add_table(RANGE("C3:D13"), options);

  workbook.save("test_table21.xlsx");
}
