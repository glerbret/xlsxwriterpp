/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outline levels");

  const xwpp::row_col_options_t level1{.hidden_ = false, .level_ = 1, .collapsed_ = false};
  const xwpp::row_col_options_t level2{.hidden_ = false, .level_ = 2, .collapsed_ = false};
  const xwpp::row_col_options_t level3{.hidden_ = false, .level_ = 3, .collapsed_ = false};
  const xwpp::row_col_options_t level4{.hidden_ = false, .level_ = 4, .collapsed_ = false};
  const xwpp::row_col_options_t level5{.hidden_ = false, .level_ = 5, .collapsed_ = false};
  const xwpp::row_col_options_t level6{.hidden_ = false, .level_ = 6, .collapsed_ = false};
  const xwpp::row_col_options_t level7{.hidden_ = false, .level_ = 7, .collapsed_ = false};

  worksheet.write_string(0, 0, "Level 1");
  worksheet.write_string(1, 0, "Level 2");
  worksheet.write_string(2, 0, "Level 3");
  worksheet.write_string(3, 0, "Level 4");
  worksheet.write_string(4, 0, "Level 5");
  worksheet.write_string(5, 0, "Level 6");
  worksheet.write_string(6, 0, "Level 7");
  worksheet.write_string(7, 0, "Level 6");
  worksheet.write_string(8, 0, "Level 5");
  worksheet.write_string(9, 0, "Level 4");
  worksheet.write_string(10, 0, "Level 3");
  worksheet.write_string(11, 0, "Level 2");
  worksheet.write_string(12, 0, "Level 1");

  worksheet.set_row(0, xwpp::DEF_ROW_HEIGHT, nullptr, level1);
  worksheet.set_row(1, xwpp::DEF_ROW_HEIGHT, nullptr, level2);
  worksheet.set_row(2, xwpp::DEF_ROW_HEIGHT, nullptr, level3);
  worksheet.set_row(3, xwpp::DEF_ROW_HEIGHT, nullptr, level4);
  worksheet.set_row(4, xwpp::DEF_ROW_HEIGHT, nullptr, level5);
  worksheet.set_row(5, xwpp::DEF_ROW_HEIGHT, nullptr, level6);
  worksheet.set_row(6, xwpp::DEF_ROW_HEIGHT, nullptr, level7);
  worksheet.set_row(7, xwpp::DEF_ROW_HEIGHT, nullptr, level6);
  worksheet.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, level5);
  worksheet.set_row(9, xwpp::DEF_ROW_HEIGHT, nullptr, level4);
  worksheet.set_row(10, xwpp::DEF_ROW_HEIGHT, nullptr, level3);
  worksheet.set_row(11, xwpp::DEF_ROW_HEIGHT, nullptr, level2);
  worksheet.set_row(12, xwpp::DEF_ROW_HEIGHT, nullptr, level1);

  workbook.save("test_outline04.xlsx");
}
