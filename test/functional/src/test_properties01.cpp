/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  const xwpp::doc_properties_t properties{
    .title        = "This is an example spreadsheet",
    .subject      = "With document properties",
    .author       = "Someone",
    .modif_author = "Someone",
    .manager      = "Dr. Heinz Doofenshmirtz",
    .company      = "of Wolves",
    .category     = "Example spreadsheets",
    .keywords     = "Sample, Example, Properties",
    .comments     = "Created with Perl and Excel::Writer::XLSX",
    .status       = "Quo",
  };

  workbook.set_properties(properties);

  worksheet.set_column(0, 0, 70);
  worksheet.write("A1", "Select 'Office Button -> Prepare -> Properties' to see the file properties.");

  workbook.save("test_properties01.xlsx");
}
