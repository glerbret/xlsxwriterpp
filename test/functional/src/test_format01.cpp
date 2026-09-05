/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  try
  {
    xwpp::workbook_t workbook;
    xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
    /* xwpp::worksheet_t& worksheet2 = */ workbook.add_worksheet("Data Sheet");
    xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();

    /* xwpp::format_t* unused1 = */ (void)workbook.add_format();
    xwpp::format_t* format = workbook.add_format();
    /* xwpp::format_t* unused2 = */ (void)workbook.add_format();
    /* xwpp::format_t* unused3 = */ (void)workbook.add_format();

    format->set_bold();

    worksheet1.write_string(0, 0, "Foo");
    worksheet1.write_number(1, 0, 123);

    worksheet3.write_string(1, 1, "Foo");
    worksheet3.write_string(2, 1, "Bar", format);
    worksheet3.write_number(3, 2, 234);

    // For testing. This doesn't have a string or format and should be ignored.
    worksheet1.write_string(0, 0, "");

    // For testing. This doesn't have a formula and should throw exception.
    worksheet1.write_formula(0, 0, "");

    workbook.save("test_format01.xlsx");
  } catch(const xwpp::xwpp_exception_t&)
  {
    return EXIT_FAILURE;
  }
}
