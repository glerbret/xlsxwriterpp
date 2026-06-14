/*
 * A simple program to write some data to an Excel file using the
 * Xlsxwriter++ library.
 *
 * This program is shown, with explanations, in Tutorial 3 of the
 * Xlsxwriter++ documentation.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

#include <chrono>

using namespace std::literals::chrono_literals;

/* Some data we want to write to the worksheet. */
struct expense
{
  char item[32];
  int cost;
  std::chrono::system_clock::time_point datetime;
};

struct expense expenses[] = {
    {"Rent", 1000, std::chrono::sys_days{2013y / std::chrono::January / 13d} + 0h + 0min + 0s},
    {"Gas",  100,  std::chrono::sys_days{2013y / std::chrono::January / 14d} + 0h + 0min + 0s},
    {"Food", 300,  std::chrono::sys_days{2013y / std::chrono::January / 16d} + 0h + 0min + 0s},
    {"Gym",  50,   std::chrono::sys_days{2013y / std::chrono::January / 20d} + 0h + 0min + 0s},
};

int main()
{
  // Create a workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  int row = 0;
  int col = 0;
  int i;

  // Add a bold format to use to highlight cells.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Add a number format for cells with money.
  xwpp::format_t* money = workbook.add_format();
  money->set_num_format("$#,##0");

  // Add an Excel date format.
  xwpp::format_t* date_format = workbook.add_format();
  date_format->set_num_format("mmmm d yyyy");

  // Adjust the column width.
  worksheet.set_column(0, 0, 15);

  // Write some data header.
  worksheet.write_string(row, col, "Item", bold);
  worksheet.write_string(row, col + 1, "Cost", bold);

  // Iterate over the data and write it out element by element.
  for(i = 0; i < 4; i++)
  {
    // Write from the first cell below the headers.
    row = i + 1;
    worksheet.write_string(row, col, expenses[i].item);
    worksheet.write_datetime(row, col + 1, expenses[i].datetime, date_format);
    worksheet.write_number(row, col + 2, expenses[i].cost, money);
  }

  // Write a total using a formula.
  worksheet.write_string(row + 1, col, "Total", bold);
  worksheet.write_formula(row + 1, col + 2, "=SUM(C2:C5)", money);

  workbook.save("tutorial03.xlsx");
}
