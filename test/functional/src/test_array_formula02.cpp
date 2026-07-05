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
  xwpp::format_t* bold         = workbook.add_format();

  bold->set_bold();

  worksheet.write_number(0, 1, 0);
  worksheet.write_number(1, 1, 0);
  worksheet.write_number(2, 1, 0);
  worksheet.write_number(0, 2, 0);
  worksheet.write_number(1, 2, 0);
  worksheet.write_number(2, 2, 0);

  worksheet.write_array_formula(RANGE("A1:A3"), "{=SUM(B1:C1*B2:C2)}", bold);

  workbook.save("test_array_formula02.xlsx");
}
