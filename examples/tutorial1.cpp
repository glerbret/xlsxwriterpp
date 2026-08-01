/*
 * A simple program to write some data to an Excel file using the
 * Xlsxwriter++ library.
 *
 * This program is shown, with explanations, in Tutorial 1 of the
 * Xlsxwriter++ documentation.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

struct expense
{
  char item[32];
  int cost;
};

struct expense expenses[] = {
    {"Rent", 1000},
    {"Gas",  100 },
    {"Food", 300 },
    {"Gym",  50  },
};

int main()
{
  // Create a workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Start from the first cell. Rows and columns are zero indexed.
  xwpp::row_num_t row = 0;
  xwpp::col_num_t col = 0;

  // Iterate over the data and write it out element by element.
  for(row = 0; row < 4; row++)
  {
    worksheet.write_string(row, col, expenses[row].item);
    worksheet.write_number(row, col + 1, expenses[row].cost);
  }

  // Write a total using a formula.
  worksheet.write_string(row, col, "Total");
  worksheet.write_formula(row, col + 1, "=SUM(B1:B4)");

  workbook.save("tutorial01.xlsx");
}
