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
  worksheet.write("B4", "Apples");
  worksheet.write("B5", "Pears");
  worksheet.write("B6", "Bananas");
  worksheet.write("B7", "Oranges");

  worksheet.write("C4", 10000, format);
  worksheet.write("C5", 2000, format);
  worksheet.write("C6", 6000, format);
  worksheet.write("C7", 500, format);

  worksheet.write("D4", 5000, format);
  worksheet.write("D5", 3000, format);
  worksheet.write("D6", 6000, format);
  worksheet.write("D7", 300, format);

  worksheet.write("E4", 8000, format);
  worksheet.write("E5", 4000, format);
  worksheet.write("E6", 6500, format);
  worksheet.write("E7", 200, format);

  worksheet.write("F4", 6000, format);
  worksheet.write("F5", 5000, format);
  worksheet.write("F6", 6000, format);
  worksheet.write("F7", 700, format);
}

}

int main()
{
  xwpp::workbook_t workbook;
  const xwpp::format_t* currency_format = workbook.format_builder().num_format("$#,##0").build();

  // Example 1. Default table with no data
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Default table with no data.");

    // Add a table to the worksheet.
    worksheet.add_table("B3:F7");
  }

  // Example 2. Default table with data
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Default table with data.");

    // Add a table to the worksheet.
    worksheet.add_table("B3:F7");

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 3. Table without default autofilter
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table without default autofilter.");

    // Set the table options.
    const xwpp::table_options_t options{.no_autofilter = true};

    // Add a table to the worksheet.
    worksheet.add_table("B3:F7", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 4. Table without default header row
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table without default header row.");

    // Set the table options.
    const xwpp::table_options_t options{.no_header_row = true};

    // Add a table to the worksheet.
    worksheet.add_table("B4:F7", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 5. Default table with "First Column" and "Last Column" options
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", R"(Default table with "First Column" and "Last Column" options.)");

    // Set the table options.
    const xwpp::table_options_t options{
      .first_column = true,
      .last_column  = true,
    };

    // Add a table to the worksheet.
    worksheet.add_table("B3:F7", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 6. Table with banded columns but without default banded rows
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table with banded columns but without default banded rows.");

    // Set the table options.
    const xwpp::table_options_t options{.no_banded_rows = true, .banded_columns = true};

    // Add a table to the worksheet.
    worksheet.add_table("B3:F7", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 7. Table with user defined column headers
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table with user defined column headers.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header = "Product"},   {.header = "Quarter 1"}, {.header = "Quarter 2"},
      {.header = "Quarter 3"}, {.header = "Quarter 4"},
    };
    const xwpp::table_options_t options{.columns = columns};

    // Add a table to the worksheet.
    worksheet.add_table("B3:F7", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 8. Table with user defined column headers
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table with user defined column headers.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header = "Product"},
      {.header = "Quarter 1"},
      {.header = "Quarter 2"},
      {.header = "Quarter 3"},
      {.header = "Quarter 4"},
      {.header = "Year", .formula = "=SUM(Table8[@[Quarter 1]:[Quarter 4]])"},
    };
    const xwpp::table_options_t options{.columns = columns};

    // Add a table to the worksheet.
    worksheet.add_table("B3:G7", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 9. Table with totals row (but no caption or totals)
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table with totals row (but no caption or totals).");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header = "Product"},
      {.header = "Quarter 1"},
      {.header = "Quarter 2"},
      {.header = "Quarter 3"},
      {.header = "Quarter 4"},
      {.header = "Year", .formula = "=SUM(Table9[@[Quarter 1]:[Quarter 4]])"},
    };
    const xwpp::table_options_t options{.total_row = true, .columns = columns};

    // Add a table to the worksheet.
    worksheet.add_table("B3:G8", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 10. Table with totals row with user captions and functions
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table with totals row with user captions and functions.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header = "Product", .total_string = "Totals"},
      {.header = "Quarter 1", .total_function = xwpp::table_total_functions_t::SUM},
      {.header = "Quarter 2", .total_function = xwpp::table_total_functions_t::SUM},
      {.header = "Quarter 3", .total_function = xwpp::table_total_functions_t::SUM},
      {.header = "Quarter 4", .total_function = xwpp::table_total_functions_t::SUM},
      {.header         = "Year",
       .formula        = "=SUM(Table10[@[Quarter 1]:[Quarter 4]])",
       .total_function = xwpp::table_total_functions_t::SUM},
    };
    const xwpp::table_options_t options{.total_row = true, .columns = columns};

    // Add a table to the worksheet.
    worksheet.add_table("B3:G8", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 11. Table with alternative Excel style
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table with alternative Excel style.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header = "Product", .total_string = "Totals"},
      {.header = "Quarter 1", .total_function = xwpp::table_total_functions_t::SUM},
      {.header = "Quarter 2", .total_function = xwpp::table_total_functions_t::SUM},
      {.header = "Quarter 3", .total_function = xwpp::table_total_functions_t::SUM},
      {.header = "Quarter 4", .total_function = xwpp::table_total_functions_t::SUM},
      {.header         = "Year",
       .formula        = "=SUM(Table11[@[Quarter 1]:[Quarter 4]])",
       .total_function = xwpp::table_total_functions_t::SUM},
    };
    const xwpp::table_options_t options{
      .style_type        = xwpp::table_style_type_t::LIGHT,
      .style_type_number = 11,
      .total_row         = true,
      .columns           = columns,
    };

    // Add a table to the worksheet.
    worksheet.add_table("B3:G8", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 12. Table with Excel style removed
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table with Excel style removed.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header = "Product", .total_string = "Totals"},
      {.header = "Quarter 1", .total_function = xwpp::table_total_functions_t::SUM},
      {.header = "Quarter 2", .total_function = xwpp::table_total_functions_t::SUM},
      {.header = "Quarter 3", .total_function = xwpp::table_total_functions_t::SUM},
      {.header = "Quarter 4", .total_function = xwpp::table_total_functions_t::SUM},
      {.header         = "Year",
       .formula        = "=SUM(Table12[@[Quarter 1]:[Quarter 4]])",
       .total_function = xwpp::table_total_functions_t::SUM},
    };
    const xwpp::table_options_t options{
      .style_type        = xwpp::table_style_type_t::LIGHT,
      .style_type_number = 0,
      .total_row         = true,
      .columns           = columns,
    };

    // Add a table to the worksheet.
    worksheet.add_table("B3:G8", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, nullptr);
  }

  // Example 13. Table with column formats
  // Set the columns widths for clarity.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column("B:G", 12);

    // Write the worksheet caption to explain the example.
    worksheet.write("B1", "Table with column formats.");

    // Set the table options.
    const std::vector<xwpp::table_column_t> columns{
      {.header = "Product", .total_string = "Totals"},
      {.header = "Quarter 1", .total_function = xwpp::table_total_functions_t::SUM, .format = currency_format},
      {.header = "Quarter 2", .total_function = xwpp::table_total_functions_t::SUM, .format = currency_format},
      {.header = "Quarter 3", .total_function = xwpp::table_total_functions_t::SUM, .format = currency_format},
      {.header = "Quarter 4", .total_function = xwpp::table_total_functions_t::SUM, .format = currency_format},
      {.header         = "Year",
       .formula        = "=SUM(Table13[@[Quarter 1]:[Quarter 4]])",
       .total_function = xwpp::table_total_functions_t::SUM,
       .format         = currency_format},
    };
    const xwpp::table_options_t options{.total_row = true, .columns = columns};

    // Add a table to the worksheet.
    worksheet.add_table("B3:G8", options);

    // Write the data into the worksheet cells.
    write_worksheet_data(worksheet, currency_format);
  }

  workbook.save("tables.xlsx");
}
