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
  xwpp::format_t* wrap         = workbook.add_format();

  wrap->set_text_wrap();

  worksheet.set_column(COLS("C:F"), 10.288);
  worksheet.set_row(2, 39);

  std::vector<xwpp::table_column_t> columns = {
      {},
      {},
      {},
      {.header_ = "Column\n4", .header_format_ = wrap}
  };
  xwpp::table_options_t options = {.columns_ = columns};

  worksheet.add_table(RANGE("C3:F13"), options);

  worksheet.write_string(CELL("A16"), "hello");

  workbook.save("test_table18.xlsx");
}
