/*
 * Example of writing dates and times in Excel using an lxw_datetime struct
 * and date formatting.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

#include <chrono>

using namespace std::literals::chrono_literals;

int main()
{
  // A datetime to display.
  const auto datetime = std::chrono::sys_days{2013y / std::chrono::February / 28d} + 12h + 0min + 0s;

  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a format with date formatting.
  xwpp::format_t* format = workbook.add_format();
  format->set_num_format("mmm d yyyy hh:mm AM/PM");

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 22);

  // Write the datetime without formatting.
  worksheet.write_datetime(0, 0, datetime); // 41333.5

  // Write the datetime with formatting.
  worksheet.write_datetime(1, 0, datetime, format); // Feb 28 2013 12:00 PM

  workbook.save("date_and_times02.xlsx");
}
