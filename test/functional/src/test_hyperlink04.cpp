/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet2 = */ workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet3 = */ workbook.add_worksheet("Data Sheet");

  workbook.unset_default_url_format();

  worksheet1.write_url(CELL("A1"), "internal:Sheet2!A1", nullptr, "", "");
  worksheet1.write_url(CELL("A3"), "internal:Sheet2!A1:A5", nullptr, "", "");
  worksheet1.write_url(CELL("A5"), "internal:'Data Sheet'!D5", nullptr, "Some text", "");
  worksheet1.write_url(CELL("E12"), "internal:Sheet1!J1", nullptr, "", "");
  worksheet1.write_url(CELL("G17"), "internal:Sheet2!A1", nullptr, "Some text", "");
  worksheet1.write_url(CELL("A18"), "internal:Sheet2!A1", nullptr, "", "Tool Tip 1");
  worksheet1.write_url(CELL("A20"), "internal:Sheet2!A1", nullptr, "More text", "Tool Tip 2");

  workbook.save("test_hyperlink04.xlsx");
}
