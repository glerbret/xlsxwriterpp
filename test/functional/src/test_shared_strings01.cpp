/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include "string"

using namespace std::string_literals;

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  std::string c = "\0"s;

  worksheet.write_string(0, 0, "_x0000_");
  for(xwpp::row_num_t i = 1; i < 127; i++)
  {
    c[0]++;
    if(i != 34)
    {
      worksheet.write_string(i, 0, c);
    }
  }

  workbook.save("test_shared_strings01.xlsx");
}
