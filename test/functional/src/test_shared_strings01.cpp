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

  char c[] = {0x00, 0x00};

  worksheet.write_string(0, 0, "_x0000_");
  for(size_t i = 1; i < 127; i++)
  {
    (*c)++;
    if(i != 34)
    {
      worksheet.write_string(i, 0, c);
    }
  }

  workbook.save("test_shared_strings01.xlsx");
}
