/*
 * Example of setting document properties such as Author, Title, etc., for an
 * Excel spreadsheet using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Create a properties structure and set some of the fields.
  const xwpp::doc_properties_t properties = {
      .title_    = "This is an example spreadsheet",
      .subject_  = "With document properties",
      .author_   = "Grégory Lerbret",
      .manager_  = "John Doe",
      .company_  = "Mine",
      .category_ = "Example spreadsheets",
      .keywords_ = "Sample, Example, Properties",
      .comments_ = "Created with Xlsxwriter++",
      .status_   = "Quo",
  };

  // Set the properties in the workbook.
  workbook.set_properties(properties);

  // Add some text to the file.
  worksheet.set_column(0, 0, 50);
  worksheet.write_string(0, 0, "Select 'Workbook Properties' to see properties.");

  workbook.save("doc_properties.xlsx");
}
