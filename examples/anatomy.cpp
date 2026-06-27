/*
 * Anatomy of a simple Xlsxwriter++ program.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook.
  xwpp::workbook_t workbook;

  // Add a worksheet with a user defined sheet name.
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet("Demo");

  // Add a worksheet with Excel's default sheet name: Sheet2.
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();

  // Add some cell formats.
  xwpp::format_t* myformat1 = workbook.add_format();
  xwpp::format_t* myformat2 = workbook.add_format();

  // Set the bold property for the first format.
  myformat1->set_bold();

  // Set a number format for the second format.
  myformat2->set_num_format("$#,##0.00");

  // Widen the first column to make the text clearer.
  worksheet1.set_column(0, 0, 20);

  // Write some unformatted data.
  worksheet1.write_string(0, 0, "Peach");
  worksheet1.write_string(1, 0, "Plum");

  // Write formatted data.
  worksheet1.write_string(2, 0, "Pear", myformat1);

  // Formats can be reused.
  worksheet1.write_string(3, 0, "Persimmon", myformat1);

  // Write some numbers.
  worksheet1.write_number(5, 0, 123);
  worksheet1.write_number(6, 0, 4567.555, myformat2);

  // Write to the second worksheet.
  worksheet2.write_string(0, 0, "Some text", myformat1);

  workbook.save("anatomy.xlsx");
}
