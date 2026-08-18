/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/chartsheet.h"
#include "xwpp/workbook.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(chartsheet)

BOOST_AUTO_TEST_CASE(assemble_xml_file)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr/>"
        "<sheetViews>"
          "<sheetView workbookViewId=\"0\"/>"
        "</sheetViews>"
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "</chartsheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();

  BOOST_CHECK_EQUAL(expected, chartsheet.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
