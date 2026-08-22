/*
 * An example of how to add conditional formatting to an Xlsxwriter++ file.
 *
 * Conditional formatting allows you to apply a format to a cell or a
 * range of cells based on certain criteria.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <vector>

namespace
{

void write_worksheet_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* format)
{
  worksheet.write_string(CELL("B4"), "Apples");
  worksheet.write_string(CELL("B5"), "Pears");
  worksheet.write_string(CELL("B6"), "Bananas");
  worksheet.write_string(CELL("B7"), "Oranges");

  worksheet.write_number(CELL("C4"), 10000, format);
  worksheet.write_number(CELL("C5"), 2000, format);
  worksheet.write_number(CELL("C6"), 6000, format);
  worksheet.write_number(CELL("C7"), 500, format);

  worksheet.write_number(CELL("D4"), 5000, format);
  worksheet.write_number(CELL("D5"), 3000, format);
  worksheet.write_number(CELL("D6"), 6000, format);
  worksheet.write_number(CELL("D7"), 300, format);

  worksheet.write_number(CELL("E4"), 8000, format);
  worksheet.write_number(CELL("E5"), 4000, format);
  worksheet.write_number(CELL("E6"), 6500, format);
  worksheet.write_number(CELL("E7"), 200, format);

  worksheet.write_number(CELL("F4"), 6000, format);
  worksheet.write_number(CELL("F5"), 5000, format);
  worksheet.write_number(CELL("F6"), 6000, format);
  worksheet.write_number(CELL("F7"), 700, format);
}

}

int main()
{
  xwpp::workbook_t workbook;

  xwpp::format_t* currency_format = workbook.add_format();
  currency_format->set_num_format("$#,##0");

  // Example 1. Default table with no data
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Default table with no data.");

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:F7"));
  }

  // Example 2. Default table with data
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Default table with data.");

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:F7"));

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 3. Table without default autofilter
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table without default autofilter.");

    // Set the table options.
    const xwpp::table_options_t options{.no_autofilter_ = true};

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:F7"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 4. Table without default header row
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table without default header row.");

    // Set the table options.
    const xwpp::table_options_t options{.no_header_row_ = true};

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B4:F7"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 5. Default table with "First Column" and "Last Column" options
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), R"(Default table with "First Column" and "Last Column" options.)");

    // Set the table options.
    const xwpp::table_options_t options{
      .first_column_ = true,
      .last_column_  = true,
    };

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:F7"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 6. Table with banded columns but without default banded rows
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table with banded columns but without default banded rows.");

    // Set the table options.
    const xwpp::table_options_t options{.no_banded_rows_ = true, .banded_columns_ = true};

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:F7"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 7. Table with user defined column headers
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table with user defined column headers.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header_ = "Product"},   {.header_ = "Quarter 1"}, {.header_ = "Quarter 2"},
      {.header_ = "Quarter 3"}, {.header_ = "Quarter 4"},
    };
    const xwpp::table_options_t options{.columns_ = columns};

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:F7"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 8. Table with user defined column headers
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table with user defined column headers.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header_ = "Product"},
      {.header_ = "Quarter 1"},
      {.header_ = "Quarter 2"},
      {.header_ = "Quarter 3"},
      {.header_ = "Quarter 4"},
      {.header_ = "Year", .formula_ = "=SUM(Table8[@[Quarter 1]:[Quarter 4]])"},
    };
    const xwpp::table_options_t options{.columns_ = columns};

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:G7"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 9. Table with totals row (but no caption or totals)
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table with totals row (but no caption or totals).");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header_ = "Product"},
      {.header_ = "Quarter 1"},
      {.header_ = "Quarter 2"},
      {
       .header_ = "Quarter 3",
       },
      {.header_ = "Quarter 4"},
      {.header_ = "Year", .formula_ = "=SUM(Table9[@[Quarter 1]:[Quarter 4]])"},
    };
    const xwpp::table_options_t options{.total_row_ = true, .columns_ = columns};

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:G8"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 10. Table with totals row with user captions and functions
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table with totals row with user captions and functions.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header_ = "Product", .total_string_ = "Totals"},
      {.header_ = "Quarter 1", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_ = "Quarter 2", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_ = "Quarter 3", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_ = "Quarter 4", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_         = "Year",
       .formula_        = "=SUM(Table10[@[Quarter 1]:[Quarter 4]])",
       .total_function_ = xwpp::table_total_functions_t::SUM},
    };
    const xwpp::table_options_t options{.total_row_ = true, .columns_ = columns};

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:G8"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 11. Table with alternative Excel style
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table with alternative Excel style.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header_ = "Product", .total_string_ = "Totals"},
      {.header_ = "Quarter 1", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_ = "Quarter 2", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_ = "Quarter 3", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_ = "Quarter 4", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_         = "Year",
       .formula_        = "=SUM(Table11[@[Quarter 1]:[Quarter 4]])",
       .total_function_ = xwpp::table_total_functions_t::SUM},
    };
    const xwpp::table_options_t options{
      .style_type_        = xwpp::table_style_type_t::LIGHT,
      .style_type_number_ = 11,
      .total_row_         = true,
      .columns_           = columns,
    };

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:G8"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 12. Table with Excel style removed
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table with Excel style removed.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header_ = "Product", .total_string_ = "Totals"},
      {.header_ = "Quarter 1", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_ = "Quarter 2", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_ = "Quarter 3", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_ = "Quarter 4", .total_function_ = xwpp::table_total_functions_t::SUM},
      {.header_         = "Year",
       .formula_        = "=SUM(Table12[@[Quarter 1]:[Quarter 4]])",
       .total_function_ = xwpp::table_total_functions_t::SUM},
    };
    const xwpp::table_options_t options{
      .style_type_        = xwpp::table_style_type_t::LIGHT,
      .style_type_number_ = 0,
      .total_row_         = true,
      .columns_           = columns,
    };

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:G8"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 13. Table with column formats
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(COLS("B:G"), 12);

    // Write the worksheet caption to explain the example.
    worksheet.write_string(CELL("B1"), "Table with column formats.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header_ = "Product", .total_string_ = "Totals"},
      {.header_ = "Quarter 1", .total_function_ = xwpp::table_total_functions_t::SUM, .format_ = currency_format},
      {.header_ = "Quarter 2", .total_function_ = xwpp::table_total_functions_t::SUM, .format_ = currency_format},
      {.header_ = "Quarter 3", .total_function_ = xwpp::table_total_functions_t::SUM, .format_ = currency_format},
      {.header_ = "Quarter 4", .total_function_ = xwpp::table_total_functions_t::SUM, .format_ = currency_format},
      {.header_         = "Year",
       .formula_        = "=SUM(Table13[@[Quarter 1]:[Quarter 4]])",
       .total_function_ = xwpp::table_total_functions_t::SUM,
       .format_         = currency_format},
    };
    const xwpp::table_options_t options{.total_row_ = true, .columns_ = columns};

    // Add a table to the worksheet.
    worksheet.add_table(RANGE("B3:G8"), options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, currency_format);
  }

  workbook.save("tables.xlsx");
}
