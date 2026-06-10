/*
 * Example of setting document properties such as Author, Title, etc., for an
 * Excel spreadsheet using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

#include <chrono>

using namespace std::literals::chrono_literals;

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  auto datetime = 2016y / std::chrono::December / 12d;

  // Set some custom document properties in the workbook.
  workbook.set_custom_property("Checked by", "Eve");
  workbook.set_custom_property("Date completed", datetime);
  workbook.set_custom_property("Document number", 12345);
  workbook.set_custom_property("Reference number", 1.2345);
  workbook.set_custom_property("Has Review", true);
  workbook.set_custom_property("Signed off", false);

  // Add some text to the file.
  worksheet.set_column(0, 0, 50);
  worksheet.write_string(0, 0, "Select 'Workbook Properties' to see properties.");

  workbook.save("doc_custom_properties.xlsx");
}
