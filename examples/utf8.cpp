/*
 * A simple Unicode UTF-8 example using Xlsxwriter++.
 *
 * Note: The source file must be UTF-8 encoded.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.write_string(2, 1, "Это фраза на русском!");

  workbook.save("utf8.xlsx");
}
