/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <chrono>

using namespace std::literals::chrono_literals;

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  workbook.set_custom_property("Checked by", "Adam");
  workbook.set_custom_property("Date completed",
                               std::chrono::sys_days{2016y / std::chrono::December / 12d} + 23h + 0min + 0s);
  workbook.set_custom_property("Document number", 12345);
  workbook.set_custom_property("Reference", 1.2345);
  workbook.set_custom_property("Source", true);
  workbook.set_custom_property("Status", false);
  workbook.set_custom_property("Department", "Finance");
  workbook.set_custom_property("Group", 1.2345678901234);

  worksheet.set_column(0, 0, 70);
  worksheet.write_string(CELL("A1"), "Select 'Office Button -> Prepare -> Properties' to see the file properties.",
                         nullptr);

  workbook.save("test_properties04.xlsx");
}
