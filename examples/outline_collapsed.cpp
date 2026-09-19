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

namespace
{

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 1-4.
void create_row_example_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* bold)
{
  // Set the column width for clarity.
  worksheet.set_column("A", 20);

  // Add data and formulas to the worksheet.
  worksheet.write("A1", "Region", bold);
  worksheet.write("A2", "North");
  worksheet.write("A3", "North");
  worksheet.write("A4", "North");
  worksheet.write("A5", "North");
  worksheet.write("A6", "North Total", bold);

  worksheet.write("B1", "Sales", bold);
  worksheet.write("B2", 1000);
  worksheet.write("B3", 1200);
  worksheet.write("B4", 900);
  worksheet.write("B5", 1200);
  worksheet.write_formula("B6", "=SUBTOTAL(9,B2:B5)", bold);

  worksheet.write("A7", "South");
  worksheet.write("A8", "South");
  worksheet.write("A9", "South");
  worksheet.write("A10", "South");
  worksheet.write("A11", "South Total", bold);

  worksheet.write("B7", 400);
  worksheet.write("B8", 600);
  worksheet.write("B9", 500);
  worksheet.write("B10", 600);
  worksheet.write_formula("B11", "=SUBTOTAL(9,B7:B10)", bold);

  worksheet.write("A12", "Grand Total", bold);
  worksheet.write_formula("B12", "=SUBTOTAL(9,B2:B10)", bold);
}

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 5-6.
void create_col_example_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* bold)
{
  // Add data and formulas to the worksheet.
  worksheet.write("A1", "Month");
  worksheet.write("B1", "Jan");
  worksheet.write("C1", "Feb");
  worksheet.write("D1", "Mar");
  worksheet.write("E1", "Apr");
  worksheet.write("F1", "May");
  worksheet.write("G1", "Jun");
  worksheet.write("H1", "Total");

  worksheet.write("A2", "North");
  worksheet.write("B2", 50);
  worksheet.write("C2", 20);
  worksheet.write("D2", 15);
  worksheet.write("E2", 25);
  worksheet.write("F2", 65);
  worksheet.write("G2", 80);
  worksheet.write_formula("H2", "=SUM(B2:G2)");

  worksheet.write("A3", "South");
  worksheet.write("B3", 10);
  worksheet.write("C3", 20);
  worksheet.write("D3", 30);
  worksheet.write("E3", 50);
  worksheet.write("F3", 50);
  worksheet.write("G3", 50);
  worksheet.write_formula("H3", "=SUM(B3:G3)");

  worksheet.write("A4", "East");
  worksheet.write("B4", 45);
  worksheet.write("C4", 75);
  worksheet.write("D4", 50);
  worksheet.write("E4", 15);
  worksheet.write("F4", 75);
  worksheet.write("G4", 100);
  worksheet.write_formula("H4", "=SUM(B4:G4)");

  worksheet.write("A5", "West");
  worksheet.write("B5", 15);
  worksheet.write("C5", 15);
  worksheet.write("D5", 55);
  worksheet.write("E5", 35);
  worksheet.write("F5", 20);
  worksheet.write("G5", 50);
  worksheet.write_formula("H5", "=SUM(B5:G5)");

  worksheet.write_formula("H6", "=SUM(H2:H5)", bold);
}

}

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

    // Set the row outline properties set.
    worksheet.set_row(1, options1);
    worksheet.set_row(2, options1);
    worksheet.set_row(3, options1);
    worksheet.set_row(4, options1);
    worksheet.set_row(5, options2);

    worksheet.set_row(6, options1);
    worksheet.set_row(7, options1);
    worksheet.set_row(8, options1);
    worksheet.set_row(9, options1);
    worksheet.set_row(10, options2);

    // Write the sub-total data that is common to the row examples.
    create_row_example_data(worksheet, bold);
  }

  // Example 2: Create a worksheet with collapsed outlined rows.
  // This is the same as the example 1  except that the all rows are collapsed.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

    // The option structs with the outline properties set.
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

    // Set the row options with the outline level.
    worksheet.set_row(1, options1);
    worksheet.set_row(2, options1);
    worksheet.set_row(3, options1);
    worksheet.set_row(4, options1);
    worksheet.set_row(5, options2);

    worksheet.set_row(6, options1);
    worksheet.set_row(7, options1);
    worksheet.set_row(8, options1);
    worksheet.set_row(9, options1);
    worksheet.set_row(10, options2);
    worksheet.set_row(11, options3);

    // Write the sub-total data that is common to the row examples.
    create_row_example_data(worksheet, bold);
  }

  // Example 3: Create a worksheet with collapsed outlined rows. Same as the
  // example 1 except that the two sub-totals are collapsed.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

    const xwpp::row_col_options_t options1{
      .hidden_    = true,
      .level_     = 2,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t options2{
      .hidden_    = false,
      .level_     = 1,
      .collapsed_ = true,
    };

    // Set the row options with the outline level.
    worksheet.set_row(1, options1);
    worksheet.set_row(2, options1);
    worksheet.set_row(3, options1);
    worksheet.set_row(4, options1);
    worksheet.set_row(5, options2);

    worksheet.set_row(6, options1);
    worksheet.set_row(7, options1);
    worksheet.set_row(8, options1);
    worksheet.set_row(9, options1);
    worksheet.set_row(10, options2);

    // Write the sub-total data that is common to the row examples.
    create_row_example_data(worksheet, bold);
  }

  // Example 4: Create a worksheet with outlined rows. Same as the example 1
  // except that the two sub-totals are collapsed.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

    const xwpp::row_col_options_t options1{
      .hidden_    = true,
      .level_     = 2,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t options2{
      .hidden_    = true,
      .level_     = 1,
      .collapsed_ = true,
    };
    const xwpp::row_col_options_t options3{
      .hidden_    = false,
      .level_     = 0,
      .collapsed_ = true,
    };

    // Set the row options with the outline level.
    worksheet.set_row(1, options1);
    worksheet.set_row(2, options1);
    worksheet.set_row(3, options1);
    worksheet.set_row(4, options1);
    worksheet.set_row(5, options2);

    worksheet.set_row(6, options1);
    worksheet.set_row(7, options1);
    worksheet.set_row(8, options1);
    worksheet.set_row(9, options1);
    worksheet.set_row(10, options2);

    worksheet.set_row(11, options3);

    // Write the sub-total data that is common to the row examples.
    create_row_example_data(worksheet, bold);
  }

  // Example 5: Create a worksheet with outlined columns.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

    const xwpp::row_col_options_t options{
      .hidden_    = false,
      .level_     = 1,
      .collapsed_ = false,
    };

    // Write the sub-total data that is common to the column examples.
    create_col_example_data(worksheet, bold);

    // Add bold format to the first row.
    worksheet.set_row(0, bold);

    // Set column formatting and the outline level.
    worksheet.set_column("A:A", 10, bold);
    worksheet.set_column("B:G", 5, nullptr, options);
    worksheet.set_column("H:H", 10);
  }

  // Example 6: Create a worksheet with outlined columns.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

    const xwpp::row_col_options_t options1{
      .hidden_    = true,
      .level_     = 1,
      .collapsed_ = false,
    };
    const xwpp::row_col_options_t options2{
      .hidden_    = false,
      .level_     = 0,
      .collapsed_ = true,
    };

    // Write the sub-total data that is common to the column examples.
    create_col_example_data(worksheet, bold);

    // Add bold format to the first row.
    worksheet.set_row(0, bold);

    // Set column formatting and the outline level.
    worksheet.set_column("A:A", 10, bold);
    worksheet.set_column("B:G", 5, nullptr, options1);
    worksheet.set_column("H:H", 10, nullptr, options2);
  }

  workbook.save("outline_collapsed.xlsx");
}
