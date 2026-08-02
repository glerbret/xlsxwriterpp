/*
 * Example of how use Xlsxwriter++ to generate Excel outlines and grouping.
 *
 * These examples focus mainly on collapsed outlines. See also the outlines.c
 * example program for more general examples.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 1-4.
void create_row_example_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* bold)
{
  // Set the column width for clarity.
  worksheet.set_column(COLS("A:A"), 20);

  // Add data and formulas to the worksheet.
  worksheet.write_string(CELL("A1"), "Region", bold);
  worksheet.write_string(CELL("A2"), "North");
  worksheet.write_string(CELL("A3"), "North");
  worksheet.write_string(CELL("A4"), "North");
  worksheet.write_string(CELL("A5"), "North");
  worksheet.write_string(CELL("A6"), "North Total", bold);

  worksheet.write_string(CELL("B1"), "Sales", bold);
  worksheet.write_number(CELL("B2"), 1000);
  worksheet.write_number(CELL("B3"), 1200);
  worksheet.write_number(CELL("B4"), 900);
  worksheet.write_number(CELL("B5"), 1200);
  worksheet.write_formula(CELL("B6"), "=SUBTOTAL(9,B2:B5)", bold);

  worksheet.write_string(CELL("A7"), "South");
  worksheet.write_string(CELL("A8"), "South");
  worksheet.write_string(CELL("A9"), "South");
  worksheet.write_string(CELL("A10"), "South");
  worksheet.write_string(CELL("A11"), "South Total", bold);

  worksheet.write_number(CELL("B7"), 400);
  worksheet.write_number(CELL("B8"), 600);
  worksheet.write_number(CELL("B9"), 500);
  worksheet.write_number(CELL("B10"), 600);
  worksheet.write_formula(CELL("B11"), "=SUBTOTAL(9,B7:B10)", bold);

  worksheet.write_string(CELL("A12"), "Grand Total", bold);
  worksheet.write_formula(CELL("B12"), "=SUBTOTAL(9,B2:B10)", bold);
}

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 5-6.
void create_col_example_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* bold)
{
  // Add data and formulas to the worksheet.
  worksheet.write_string(CELL("A1"), "Month");
  worksheet.write_string(CELL("B1"), "Jan");
  worksheet.write_string(CELL("C1"), "Feb");
  worksheet.write_string(CELL("D1"), "Mar");
  worksheet.write_string(CELL("E1"), "Apr");
  worksheet.write_string(CELL("F1"), "May");
  worksheet.write_string(CELL("G1"), "Jun");
  worksheet.write_string(CELL("H1"), "Total");

  worksheet.write_string(CELL("A2"), "North");
  worksheet.write_number(CELL("B2"), 50);
  worksheet.write_number(CELL("C2"), 20);
  worksheet.write_number(CELL("D2"), 15);
  worksheet.write_number(CELL("E2"), 25);
  worksheet.write_number(CELL("F2"), 65);
  worksheet.write_number(CELL("G2"), 80);
  worksheet.write_formula(CELL("H2"), "=SUM(B2:G2)");

  worksheet.write_string(CELL("A3"), "South");
  worksheet.write_number(CELL("B3"), 10);
  worksheet.write_number(CELL("C3"), 20);
  worksheet.write_number(CELL("D3"), 30);
  worksheet.write_number(CELL("E3"), 50);
  worksheet.write_number(CELL("F3"), 50);
  worksheet.write_number(CELL("G3"), 50);
  worksheet.write_formula(CELL("H3"), "=SUM(B3:G3)");

  worksheet.write_string(CELL("A4"), "East");
  worksheet.write_number(CELL("B4"), 45);
  worksheet.write_number(CELL("C4"), 75);
  worksheet.write_number(CELL("D4"), 50);
  worksheet.write_number(CELL("E4"), 15);
  worksheet.write_number(CELL("F4"), 75);
  worksheet.write_number(CELL("G4"), 100);
  worksheet.write_formula(CELL("H4"), "=SUM(B4:G4)");

  worksheet.write_string(CELL("A5"), "West");
  worksheet.write_number(CELL("B5"), 15);
  worksheet.write_number(CELL("C5"), 15);
  worksheet.write_number(CELL("D5"), 55);
  worksheet.write_number(CELL("E5"), 35);
  worksheet.write_number(CELL("F5"), 20);
  worksheet.write_number(CELL("G5"), 50);
  worksheet.write_formula(CELL("H5"), "=SUM(B5:G5)");

  worksheet.write_formula(CELL("H6"), "=SUM(H2:H5)", bold);
}

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet("Outlined Rows");
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet("Collapsed Rows 1");
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet("Collapsed Rows 2");
  xwpp::worksheet_t& worksheet4 = workbook.add_worksheet("Collapsed Rows 3");
  xwpp::worksheet_t& worksheet5 = workbook.add_worksheet("Outline Columns");
  xwpp::worksheet_t& worksheet6 = workbook.add_worksheet("Collapsed Columns");

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  /*
   * Example 1: Create a worksheet with outlined rows. It also includes
   * SUBTOTAL() functions so that it looks like the type of automatic
   * outlines that are generated when you use the 'Sub Totals' option.
   *
   * For outlines the important parameters are 'hidden' and 'level'. Rows
   * with the same 'level' are grouped together. The group will be collapsed
   * if 'hidden' is non-zero.
   */

  // The option structs with the outline level set.
  xwpp::row_col_options_t options1{
    .hidden_    = false,
    .level_     = 2,
    .collapsed_ = false,
  };
  xwpp::row_col_options_t options2{
    .hidden_    = false,
    .level_     = 1,
    .collapsed_ = false,
  };

  // Set the row outline properties set.
  worksheet1.set_row(1, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet1.set_row(2, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet1.set_row(3, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet1.set_row(4, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet1.set_row(5, xwpp::DEF_ROW_HEIGHT, nullptr, options2);

  worksheet1.set_row(6, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet1.set_row(7, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet1.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet1.set_row(9, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet1.set_row(10, xwpp::DEF_ROW_HEIGHT, nullptr, options2);

  // Write the sub-total data that is common to the row examples.
  create_row_example_data(worksheet1, bold);

  // Example 2: Create a worksheet with collapsed outlined rows.
  // This is the same as the example 1  except that the all rows are collapsed.

  // The option structs with the outline properties set.
  xwpp::row_col_options_t options3{
    .hidden_    = true,
    .level_     = 2,
    .collapsed_ = false,
  };
  xwpp::row_col_options_t options4{
    .hidden_    = true,
    .level_     = 1,
    .collapsed_ = false,
  };
  xwpp::row_col_options_t options5{
    .hidden_    = false,
    .level_     = 0,
    .collapsed_ = true,
  };

  // Set the row options with the outline level.
  worksheet2.set_row(1, xwpp::DEF_ROW_HEIGHT, nullptr, options3);
  worksheet2.set_row(2, xwpp::DEF_ROW_HEIGHT, nullptr, options3);
  worksheet2.set_row(3, xwpp::DEF_ROW_HEIGHT, nullptr, options3);
  worksheet2.set_row(4, xwpp::DEF_ROW_HEIGHT, nullptr, options3);
  worksheet2.set_row(5, xwpp::DEF_ROW_HEIGHT, nullptr, options4);

  worksheet2.set_row(6, xwpp::DEF_ROW_HEIGHT, nullptr, options3);
  worksheet2.set_row(7, xwpp::DEF_ROW_HEIGHT, nullptr, options3);
  worksheet2.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, options3);
  worksheet2.set_row(9, xwpp::DEF_ROW_HEIGHT, nullptr, options3);
  worksheet2.set_row(10, xwpp::DEF_ROW_HEIGHT, nullptr, options4);
  worksheet2.set_row(11, xwpp::DEF_ROW_HEIGHT, nullptr, options5);

  // Write the sub-total data that is common to the row examples.
  create_row_example_data(worksheet2, bold);

  // Example 3: Create a worksheet with collapsed outlined rows. Same as the
  // example 1 except that the two sub-totals are collapsed.
  xwpp::row_col_options_t options6{
    .hidden_    = true,
    .level_     = 2,
    .collapsed_ = false,
  };
  xwpp::row_col_options_t options7{
    .hidden_    = false,
    .level_     = 1,
    .collapsed_ = true,
  };

  // Set the row options with the outline level.
  worksheet3.set_row(1, xwpp::DEF_ROW_HEIGHT, nullptr, options6);
  worksheet3.set_row(2, xwpp::DEF_ROW_HEIGHT, nullptr, options6);
  worksheet3.set_row(3, xwpp::DEF_ROW_HEIGHT, nullptr, options6);
  worksheet3.set_row(4, xwpp::DEF_ROW_HEIGHT, nullptr, options6);
  worksheet3.set_row(5, xwpp::DEF_ROW_HEIGHT, nullptr, options7);

  worksheet3.set_row(6, xwpp::DEF_ROW_HEIGHT, nullptr, options6);
  worksheet3.set_row(7, xwpp::DEF_ROW_HEIGHT, nullptr, options6);
  worksheet3.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, options6);
  worksheet3.set_row(9, xwpp::DEF_ROW_HEIGHT, nullptr, options6);
  worksheet3.set_row(10, xwpp::DEF_ROW_HEIGHT, nullptr, options7);

  // Write the sub-total data that is common to the row examples.
  create_row_example_data(worksheet3, bold);

  // Example 4: Create a worksheet with outlined rows. Same as the example 1
  // except that the two sub-totals are collapsed.
  xwpp::row_col_options_t options8{
    .hidden_    = true,
    .level_     = 2,
    .collapsed_ = false,
  };
  xwpp::row_col_options_t options9{
    .hidden_    = true,
    .level_     = 1,
    .collapsed_ = true,
  };
  xwpp::row_col_options_t options10{
    .hidden_    = false,
    .level_     = 0,
    .collapsed_ = true,
  };

  // Set the row options with the outline level.
  worksheet4.set_row(1, xwpp::DEF_ROW_HEIGHT, nullptr, options8);
  worksheet4.set_row(2, xwpp::DEF_ROW_HEIGHT, nullptr, options8);
  worksheet4.set_row(3, xwpp::DEF_ROW_HEIGHT, nullptr, options8);
  worksheet4.set_row(4, xwpp::DEF_ROW_HEIGHT, nullptr, options8);
  worksheet4.set_row(5, xwpp::DEF_ROW_HEIGHT, nullptr, options9);

  worksheet4.set_row(6, xwpp::DEF_ROW_HEIGHT, nullptr, options8);
  worksheet4.set_row(7, xwpp::DEF_ROW_HEIGHT, nullptr, options8);
  worksheet4.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, options8);
  worksheet4.set_row(9, xwpp::DEF_ROW_HEIGHT, nullptr, options8);
  worksheet4.set_row(10, xwpp::DEF_ROW_HEIGHT, nullptr, options9);

  worksheet4.set_row(11, xwpp::DEF_ROW_HEIGHT, nullptr, options10);

  // Write the sub-total data that is common to the row examples.
  create_row_example_data(worksheet4, bold);

  // Example 5: Create a worksheet with outlined columns.

  xwpp::row_col_options_t options11{
    .hidden_    = false,
    .level_     = 1,
    .collapsed_ = false,
  };

  // Write the sub-total data that is common to the column examples.
  create_col_example_data(worksheet5, bold);

  // Add bold format to the first row.
  worksheet5.set_row(0, xwpp::DEF_ROW_HEIGHT, bold, std::nullopt);

  // Set column formatting and the outline level.
  worksheet5.set_column(COLS("A:A"), 10, bold, std::nullopt);
  worksheet5.set_column(COLS("B:G"), 5, nullptr, options11);
  worksheet5.set_column(COLS("H:H"), 10);

  // Example 6: Create a worksheet with outlined columns.
  xwpp::row_col_options_t options12{
    .hidden_    = true,
    .level_     = 1,
    .collapsed_ = false,
  };
  xwpp::row_col_options_t options13{
    .hidden_    = false,
    .level_     = 0,
    .collapsed_ = true,
  };

  // Write the sub-total data that is common to the column examples.
  create_col_example_data(worksheet6, bold);

  // Add bold format to the first row.
  worksheet6.set_row(0, xwpp::DEF_ROW_HEIGHT, bold, std::nullopt);

  // Set column formatting and the outline level.
  worksheet6.set_column(COLS("A:A"), 10, bold, std::nullopt);
  worksheet6.set_column(COLS("B:G"), 5, nullptr, options12);
  worksheet6.set_column(COLS("H:H"), 10, nullptr, options13);

  workbook.save("outline_collapsed.xlsx");
}
