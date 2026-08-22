/*
 * Example of how use Xlsxwriter++ to generate Excel outlines and grouping.
 *
 * Excel allows you to group rows or columns so that they can be hidden or
 * displayed with a single mouse click. This feature is referred to as
 * outlines.
 *
 * Outlines can reduce complex data down to a few salient sub-totals or
 * summaries.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Example 1: Create a worksheet with outlined rows. It also includes
  // SUBTOTAL() functions so that it looks like the type of automatic
  // outlines that are generated when you use the 'Sub Totals' option.
  //
  // For outlines the important parameters are 'hidden' and 'level'. Rows
  // with the same 'level' are grouped together. The group will be collapsed
  // if 'hidden' is non-zero.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

    // The option structs with the outline level set.
    const xwpp::row_col_options_t options1{
      .hidden_    = false,
      .level_     = 2,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t options2{
      .hidden_    = false,
      .level_     = 1,
      .collapsed_ = false,
    };

    // Set the column width for clarity.
    worksheet.set_column(COLS("A:A"), 20);

    // Set the row options with the outline level.
    worksheet.set_row(1, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(2, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(3, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(4, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(5, xwpp::DEF_ROW_HEIGHT, nullptr, options2);

    worksheet.set_row(6, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(7, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(9, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(10, xwpp::DEF_ROW_HEIGHT, nullptr, options2);

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

  // Example 2: Create a worksheet with outlined rows. This is the same as
  // the previous example except that the rows are collapsed.  Note: We need
  // to indicate the row that contains the collapsed symbol '+' with the
  // optional parameter, 'collapsed'.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

    // The option structs with the outline level and collapsed property set.
    const xwpp::row_col_options_t options1{
      .hidden_    = true,
      .level_     = 2,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t options2{
      .hidden_    = true,
      .level_     = 1,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t options3{
      .hidden_    = false,
      .level_     = 0,
      .collapsed_ = true,
    };

    // Set the column width for clarity.
    worksheet.set_column(COLS("A:A"), 20);

    // Set the row options with the outline level.
    worksheet.set_row(1, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(2, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(3, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(4, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(5, xwpp::DEF_ROW_HEIGHT, nullptr, options2);

    worksheet.set_row(6, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(7, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(9, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
    worksheet.set_row(10, xwpp::DEF_ROW_HEIGHT, nullptr, options2);
    worksheet.set_row(11, xwpp::DEF_ROW_HEIGHT, nullptr, options3);

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

  // Example 3: Create a worksheet with outlined columns.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

    const xwpp::row_col_options_t options{
      .hidden_    = false,
      .level_     = 1,
      .collapsed_ = false,
    };

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

    // Add bold format to the first row.
    worksheet.set_row(0, xwpp::DEF_ROW_HEIGHT, bold);

    // Set column formatting and the outline level.
    worksheet.set_column(COLS("A:A"), 10, bold);
    worksheet.set_column(COLS("B:G"), 5, nullptr, options);
    worksheet.set_column(COLS("H:H"), 10);
  }

  // Example 4: Show all possible outline levels.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

    const xwpp::row_col_options_t level1{
      .hidden_    = false,
      .level_     = 1,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t level2{
      .hidden_    = false,
      .level_     = 2,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t level3{
      .hidden_    = false,
      .level_     = 3,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t level4{
      .hidden_    = false,
      .level_     = 4,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t level5{
      .hidden_    = false,
      .level_     = 5,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t level6{
      .hidden_    = false,
      .level_     = 6,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t level7{
      .hidden_    = false,
      .level_     = 7,
      .collapsed_ = false,
    };

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
  }

  workbook.save("outline.xlsx");
}
