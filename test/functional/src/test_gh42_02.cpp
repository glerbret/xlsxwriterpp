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

  //    char string[] = {0xe5, 0x9b, 0xbe, 0x20, 0xe5, 0x9b, 0xbe, 0x00};

  worksheet.write_string(0, 0, "\xe5\x9b\xbe\x20\xe5\x9b\xbe");

  workbook.save("test_gh42_02.xlsx");
}
