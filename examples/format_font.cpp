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

  // Write some formatted strings.
  {
    // Set the bold property for format 1.
    xwpp::format_t* format = workbook.add_format();
    format->set_bold();

    worksheet.write_string(0, 0, "This is bold", format);
  }

  {
    // Set the italic property for format 2.
    xwpp::format_t* format = workbook.add_format();
    format->set_italic();

    worksheet.write_string(1, 0, "This is italic", format);
  }

  {
    // Set the bold and italic properties for format 3.
    xwpp::format_t* format = workbook.add_format();
    format->set_bold();
    format->set_italic();

    worksheet.write_string(2, 0, "Bold and italic", format);
  }

  workbook.save("format_font.xlsx");
}
