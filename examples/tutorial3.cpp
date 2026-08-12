/*
 * A simple program to write some data to an Excel file using the
 * Xlsxwriter++ library.
 *
 * This program is shown, with explanations, in Tutorial 3 of the
 * Xlsxwriter++ documentation.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <chrono>
#include <string>
#include <vector>

using namespace std::literals::chrono_literals;

int main()
{
  struct expense
  {
    std::string item_;
    int cost_;
    std::chrono::system_clock::time_point datetime_;
  };

  const std::vector<expense> expenses{
    {.item_     = "Rent",
     .cost_     = 1000,
     .datetime_ = std::chrono::sys_days{2013y / std::chrono::January / 13d} + 0h + 0min + 0s},
    {.item_     = "Gas",
     .cost_     = 100,
     .datetime_ = std::chrono::sys_days{2013y / std::chrono::January / 14d} + 0h + 0min + 0s},
    {.item_     = "Food",
     .cost_     = 300,
     .datetime_ = std::chrono::sys_days{2013y / std::chrono::January / 16d} + 0h + 0min + 0s},
    {.item_     = "Gym",
     .cost_     = 50,
     .datetime_ = std::chrono::sys_days{2013y / std::chrono::January / 20d} + 0h + 0min + 0s},
  };

  // Create a workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

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

  xwpp::row_num_t row_num = 0;

  // Write some data header.
  worksheet.write_string(row_num, 0, "Item", bold);
  worksheet.write_string(row_num, 1, "Date", bold);
  worksheet.write_string(row_num, 2, "Cost", bold);
  row_num++;

  // Iterate over the data and write it out element by element.
  for(const auto& value: expenses)
  {
    worksheet.write_string(row_num, 0, value.item_);
    worksheet.write_datetime(row_num, 1, value.datetime_, date_format);
    worksheet.write_number(row_num, 2, value.cost_, money);
    row_num++;
  }

  // Write a total using a formula.
  worksheet.write_string(row_num, 0, "Total", bold);
  worksheet.write_formula(row_num, 2, "=SUM(C2:C5)", money);

  workbook.save("tutorial03.xlsx");
}
