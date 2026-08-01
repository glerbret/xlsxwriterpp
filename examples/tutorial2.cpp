/*
 * A simple program to write some data to an Excel file using the
 * Xlsxwriter++ library.
 *
 * This program is shown, with explanations, in Tutorial 2 of the
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

  xwpp::row_num_t row = 0;
  xwpp::col_num_t col = 0;
  int i;

  // Add a bold format to use to highlight cells.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Add a number format for cells with money.
  xwpp::format_t* money = workbook.add_format();
  money->set_num_format("$#,##0");

  // Write some data header.
  worksheet.write_string(row, col, "Item", bold);
  worksheet.write_string(row, col + 1, "Cost", bold);

  // Iterate over the data and write it out element by element.
  for(i = 0; i < 4; i++)
  {
    // Write from the first cell below the headers.
    row = i + 1;
    worksheet.write_string(row, col, expenses[i].item);
    worksheet.write_number(row, col + 1, expenses[i].cost, money);
  }

  // Write a total using a formula.
  worksheet.write_string(row + 1, col, "Total", bold);
  worksheet.write_formula(row + 1, col + 1, "=SUM(B2:B5)", money);

  workbook.save("tutorial02.xlsx");
}
