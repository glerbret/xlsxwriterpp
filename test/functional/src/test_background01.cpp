/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <filesystem>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  const std::filesystem::path path = std::filesystem::path{"images"} / "logo.jpg";
  worksheet.insert_image("E9", path);

  workbook.save("test_background01.xlsx");
}
