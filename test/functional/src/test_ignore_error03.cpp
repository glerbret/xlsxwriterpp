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

  for(xwpp::row_num_t row = 0; row < 10; row++)
  {
    worksheet.write_string(row, 0, "123", NULL);
  }

  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "A1:A10");

  workbook.save("test_ignore_error03.xlsx");
}
