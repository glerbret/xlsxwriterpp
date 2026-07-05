/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <chrono>

using namespace std::literals::chrono_literals;

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::format_t* format = workbook.add_format();
  format->set_num_format_index(14);

  const auto datetime1 = std::chrono::sys_days{1900y / std::chrono::January / 1d} + 0h + 0min + 0s;
  const auto datetime2 = std::chrono::sys_days{1902y / std::chrono::September / 26d} + 0h + 0min + 0s;
  const auto datetime3 = std::chrono::sys_days{1913y / std::chrono::September / 8d} + 0h + 0min + 0s;
  const auto datetime4 = std::chrono::sys_days{1927y / std::chrono::May / 18d} + 0h + 0min + 0s;
  const auto datetime5 = std::chrono::sys_days{2173y / std::chrono::October / 14d} + 0h + 0min + 0s;
  // TODO Won't work
  const auto datetime6 = std::chrono::sys_days{4637y / std::chrono::November / 26d} + 0h + 0min + 0s;

  // TODO Add again this struct
  //  lxw_datetime datetime1 = {1900, 1, 1, 0, 0, 0.0};
  //  lxw_datetime datetime2 = {1902, 9, 26, 0, 0, 0.0};
  //  lxw_datetime datetime3 = {1913, 9, 8, 0, 0, 0.0};
  //  lxw_datetime datetime4 = {1927, 5, 18, 0, 0, 0.0};
  //  lxw_datetime datetime5 = {2173, 10, 14, 0, 0, 0.0};
  //  lxw_datetime datetime6 = {4637, 11, 26, 0, 0, 0.0};

  worksheet.set_column(0, 0, 12);

  worksheet.write_datetime(CELL("A1"), datetime1, format);
  worksheet.write_datetime(CELL("A2"), datetime2, format);
  worksheet.write_datetime(CELL("A3"), datetime3, format);
  worksheet.write_datetime(CELL("A4"), datetime4, format);
  worksheet.write_datetime(CELL("A5"), datetime5, format);
  worksheet.write_datetime(CELL("A6"), datetime6, format);

  workbook.save("test_date_1904_01.xlsx");
}
