/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/workbook.h"
#include "xwpp/worksheet.h"

#include <string>

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(worksheet)

BOOST_AUTO_TEST_CASE(worksheet01)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData/>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(worksheet02)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1\" spans=\"1:1\">"
          "<c r=\"A1\">"
            "<v>123</v>"
          "</c>"
        "</row>"
      "</sheetData>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();
  worksheet.write_number(0, 0, 123);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(worksheet03)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1:E9\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1\" spans=\"1:5\">"
          "<c r=\"A1\" t=\"s\">"
            "<v>0</v>"
          "</c>"
        "</row>"
        "<row r=\"2\" spans=\"1:5\">"
          "<c r=\"C2\">"
            "<v>123</v>"
          "</c>"
        "</row>"
        "<row r=\"4\" spans=\"1:5\">"
          "<c r=\"B4\" t=\"s\">"
            "<v>1</v>"
          "</c>"
        "</row>"
        "<row r=\"9\" spans=\"1:5\">"
          "<c r=\"E9\">"
            "<v>890</v>"
          "</c>"
        "</row>"
      "</sheetData>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();
  worksheet.write_string(0, 0, "Foo");
  worksheet.write_number(1, 2, 123);
  worksheet.write_string(3, 1, "Bar");
  worksheet.write_number(8, 4, 890);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
