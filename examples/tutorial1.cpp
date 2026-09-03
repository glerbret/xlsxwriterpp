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

#include <string>
#include <vector>

int main()
{
  struct expense
  {
    std::string item_;
    int cost_{0};
  };

  const std::vector<expense> expenses{
    {.item_ = "Rent", .cost_ = 1000},
    {.item_ = "Gas",  .cost_ = 100 },
    {.item_ = "Food", .cost_ = 300 },
    {.item_ = "Gym",  .cost_ = 50  },
  };

  // Create a workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Start from the first cell.
  xwpp::row_num_t row_num{0};

  // Iterate over the data and write it out element by element.
  for(const auto& value: expenses)
  {
    worksheet.write_string(row_num, 0, value.item_);
    worksheet.write_number(row_num, 1, value.cost_);
    row_num++;
  }

  // Write a total using a formula.
  worksheet.write_string(row_num, 0, "Total");
  worksheet.write_formula(row_num, 1, "=SUM(B1:B4)");

  workbook.save("tutorial01.xlsx");
}
