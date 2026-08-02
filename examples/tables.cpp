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

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1  = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2  = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet3  = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet4  = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet5  = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet6  = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet7  = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet8  = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet9  = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet10 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet11 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet12 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet13 = workbook.add_worksheet();

  xwpp::format_t* currency_format = workbook.add_format();
  currency_format->set_num_format("$#,##0");

  // Example 1. Default table with no data
  // Set the columns widths for clarity.
  worksheet1.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet1.write_string(CELL("B1"), "Default table with no data.");

  // Add a table to the worksheet.
  worksheet1.add_table(RANGE("B3:F7"));

  // Example 2. Default table with data
  // Set the columns widths for clarity.
  worksheet2.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet2.write_string(CELL("B1"), "Default table with data.");

  // Add a table to the worksheet.
  worksheet2.add_table(RANGE("B3:F7"));

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet2, nullptr);

  // Example 3. Table without default autofilter
  // Set the columns widths for clarity.
  worksheet3.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet3.write_string(CELL("B1"), "Table without default autofilter.");

  // Set the table options.
  xwpp::table_options_t options3{.no_autofilter_ = true};

  // Add a table to the worksheet.
  worksheet3.add_table(RANGE("B3:F7"), options3);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet3, nullptr);

  // Example 4. Table without default header row
  // Set the columns widths for clarity.
  worksheet4.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet4.write_string(CELL("B1"), "Table without default header row.");

  // Set the table options.
  xwpp::table_options_t options4{.no_header_row_ = true};

  // Add a table to the worksheet.
  worksheet4.add_table(RANGE("B4:F7"), options4);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet4, nullptr);

  // Example 5. Default table with "First Column" and "Last Column" options
  // Set the columns widths for clarity.
  worksheet5.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet5.write_string(CELL("B1"), "Default table with \"First Column\" and \"Last Column\" options.");

  // Set the table options.
  xwpp::table_options_t options5{
    .first_column_ = true,
    .last_column_  = true,
  };

  // Add a table to the worksheet.
  worksheet5.add_table(RANGE("B3:F7"), options5);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet5, nullptr);

  // Example 6. Table with banded columns but without default banded rows
  // Set the columns widths for clarity.
  worksheet6.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet6.write_string(CELL("B1"), "Table with banded columns but without default banded rows.");

  // Set the table options.
  xwpp::table_options_t options6 = {.no_banded_rows_ = true, .banded_columns_ = true};

  // Add a table to the worksheet.
  worksheet6.add_table(RANGE("B3:F7"), options6);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet6, nullptr);

  // Example 7. Table with user defined column headers
  // Set the columns widths for clarity.
  worksheet7.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet7.write_string(CELL("B1"), "Table with user defined column headers.");

  // Set the table options.

  std::vector<xwpp::table_column_t> columns7{
    {.header_ = "Product"},   {.header_ = "Quarter 1"}, {.header_ = "Quarter 2"},
    {.header_ = "Quarter 3"}, {.header_ = "Quarter 4"},
  };

  xwpp::table_options_t options7 = {.columns_ = columns7};

  // Add a table to the worksheet.
  worksheet7.add_table(RANGE("B3:F7"), options7);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet7, nullptr);

  // Example 8. Table with user defined column headers
  // Set the columns widths for clarity.
  worksheet8.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet8.write_string(CELL("B1"), "Table with user defined column headers.");

  // Set the table options.
  std::vector<xwpp::table_column_t> columns8{
    {.header_ = "Product"},
    {.header_ = "Quarter 1"},
    {.header_ = "Quarter 2"},
    {.header_ = "Quarter 3"},
    {.header_ = "Quarter 4"},
    {.header_ = "Year", .formula_ = "=SUM(Table8[@[Quarter 1]:[Quarter 4]])"},
  };

  xwpp::table_options_t options8 = {.columns_ = columns8};

  // Add a table to the worksheet.
  worksheet8.add_table(RANGE("B3:G7"), options8);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet8, nullptr);

  // Example 9. Table with totals row (but no caption or totals)
  // Set the columns widths for clarity.
  worksheet9.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet9.write_string(CELL("B1"), "Table with totals row (but no caption or totals).");

  // Set the table options.
  std::vector<xwpp::table_column_t> columns9{
    {.header_ = "Product"},
    {.header_ = "Quarter 1"},
    {.header_ = "Quarter 2"},
    {
     .header_ = "Quarter 3",
     },
    {.header_ = "Quarter 4"},
    {.header_ = "Year", .formula_ = "=SUM(Table9[@[Quarter 1]:[Quarter 4]])"},
  };

  xwpp::table_options_t options9{.total_row_ = true, .columns_ = columns9};

  // Add a table to the worksheet.
  worksheet9.add_table(RANGE("B3:G8"), options9);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet9, nullptr);

  // Example 10. Table with totals row with user captions and functions
  // Set the columns widths for clarity.
  worksheet10.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet10.write_string(CELL("B1"), "Table with totals row with user captions and functions.");

  // Set the table options.
  std::vector<xwpp::table_column_t> columns10{
    {.header_ = "Product", .total_string_ = "Totals"},
    {.header_ = "Quarter 1", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Quarter 2", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Quarter 3", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Quarter 4", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_         = "Year",
     .formula_        = "=SUM(Table10[@[Quarter 1]:[Quarter 4]])",
     .total_function_ = xwpp::table_total_functions_t::SUM},
  };

  xwpp::table_options_t options10 = {.total_row_ = true, .columns_ = columns10};

  // Add a table to the worksheet.
  worksheet10.add_table(RANGE("B3:G8"), options10);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet10, nullptr);

  // Example 11. Table with alternative Excel style
  // Set the columns widths for clarity.
  worksheet11.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet11.write_string(CELL("B1"), "Table with alternative Excel style.");

  // Set the table options.
  std::vector<xwpp::table_column_t> columns11{
    {.header_ = "Product", .total_string_ = "Totals"},
    {.header_ = "Quarter 1", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Quarter 2", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Quarter 3", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Quarter 4", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_         = "Year",
     .formula_        = "=SUM(Table11[@[Quarter 1]:[Quarter 4]])",
     .total_function_ = xwpp::table_total_functions_t::SUM},
  };

  xwpp::table_options_t options11{
    .style_type_        = xwpp::table_style_type_t::LIGHT,
    .style_type_number_ = 11,
    .total_row_         = true,
    .columns_           = columns11,
  };

  // Add a table to the worksheet.
  worksheet11.add_table(RANGE("B3:G8"), options11);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet11, nullptr);

  // Example 12. Table with Excel style removed
  // Set the columns widths for clarity.
  worksheet12.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet12.write_string(CELL("B1"), "Table with Excel style removed.");

  // Set the table options.
  std::vector<xwpp::table_column_t> columns12{
    {.header_ = "Product", .total_string_ = "Totals"},
    {.header_ = "Quarter 1", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Quarter 2", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Quarter 3", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Quarter 4", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_         = "Year",
     .formula_        = "=SUM(Table12[@[Quarter 1]:[Quarter 4]])",
     .total_function_ = xwpp::table_total_functions_t::SUM},
  };

  xwpp::table_options_t options12{
    .style_type_        = xwpp::table_style_type_t::LIGHT,
    .style_type_number_ = 0,
    .total_row_         = true,
    .columns_           = columns12,
  };

  // Add a table to the worksheet.
  worksheet12.add_table(RANGE("B3:G8"), options12);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet12, nullptr);

  // Example 13. Table with column formats
  // Set the columns widths for clarity.
  worksheet13.set_column(COLS("B:G"), 12);

  // Write the worksheet caption to explain the example.
  worksheet13.write_string(CELL("B1"), "Table with column formats.");

  // Set the table options.

  std::vector<xwpp::table_column_t> columns13{
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

  xwpp::table_options_t options13{.total_row_ = true, .columns_ = columns13};

  // Add a table to the worksheet.
  worksheet13.add_table(RANGE("B3:G8"), options13);

  // Write the data into the worksheet cells.
  write_worksheet_data(worksheet13, currency_format);

  workbook.save("tables.xlsx");
}
