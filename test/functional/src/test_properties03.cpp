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

  workbook.set_custom_property("Checked by", "Adam");

  worksheet.set_column(0, 0, 70);
  worksheet.write_string(CELL("A1"), "Select 'Office Button -> Prepare -> Properties' to see the file properties.",
                         nullptr);

  workbook.save("test_properties03.xlsx");
}
