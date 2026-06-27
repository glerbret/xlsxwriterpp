/*
 * Example of writing some data with font formatting to a simple Excel
 * file using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 20);

  // Add some formats.
  xwpp::format_t* format1 = workbook.add_format();
  xwpp::format_t* format2 = workbook.add_format();
  xwpp::format_t* format3 = workbook.add_format();

  // Set the bold property for format 1.
  format1->set_bold();

  // Set the italic property for format 2.
  format2->set_italic();

  // Set the bold and italic properties for format 3.
  format3->set_bold();
  format3->set_italic();

  // Write some formatted strings.
  worksheet.write_string(0, 0, "This is bold", format1);
  worksheet.write_string(1, 0, "This is italic", format2);
  worksheet.write_string(2, 0, "Bold and italic", format3);

  workbook.save("format_font.xlsx");
}
