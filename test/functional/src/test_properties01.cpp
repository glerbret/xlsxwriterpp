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
  xwpp::doc_properties_t properties{
      .title_        = "This is an example spreadsheet",
      .subject_      = "With document properties",
      .author_       = "Someone",
      .modif_author_ = "Someone",
      .manager_      = "Dr. Heinz Doofenshmirtz",
      .company_      = "of Wolves",
      .category_     = "Example spreadsheets",
      .keywords_     = "Sample, Example, Properties",
      .comments_     = "Created with Perl and Excel::Writer::XLSX",
      .status_       = "Quo",
  };

  workbook.set_properties(properties);

  worksheet.set_column(0, 0, 70);
  worksheet.write_string(CELL("A1"), "Select 'Office Button -> Prepare -> Properties' to see the file properties.",
                         nullptr);

  workbook.save("test_properties01.xlsx");
}
