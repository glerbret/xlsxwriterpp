/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/workbook.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(workbook)

BOOST_AUTO_TEST_CASE(workbook01)
{
  const std::string expected =
      // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>"
      "<workbookPr defaultThemeVersion=\"124226\"/>"
      "<bookViews>"
        "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>"
      "</bookViews>"
      "<sheets>"
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>"
      "</sheets>"
      "<calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>"
    "</workbook>";
  // clang-format on

  xwpp::workbook_t workbook;
  workbook.add_worksheet();

  BOOST_CHECK_EQUAL(expected, workbook.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(workbook02)
{
  const std::string expected =
      // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>"
      "<workbookPr defaultThemeVersion=\"124226\"/>"
      "<bookViews>"
        "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>"
      "</bookViews>"
      "<sheets>"
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>"
        "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>"
      "</sheets>"
      "<calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>"
    "</workbook>";
  // clang-format on

  xwpp::workbook_t workbook;
  workbook.add_worksheet();
  workbook.add_worksheet();

  BOOST_CHECK_EQUAL(expected, workbook.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(workbook03)
{
  const std::string expected =
      // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>"
      "<workbookPr defaultThemeVersion=\"124226\"/>"
      "<bookViews>"
        "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>"
      "</bookViews>"
      "<sheets>"
        "<sheet name=\"Non Default Name\" sheetId=\"1\" r:id=\"rId1\"/>"
        "<sheet name=\"Another Name\" sheetId=\"2\" r:id=\"rId2\"/>"
      "</sheets>"
      "<calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>"
    "</workbook>";
  // clang-format on

  xwpp::workbook_t workbook;
  workbook.add_worksheet("Non Default Name");
  workbook.add_worksheet("Another Name");

  BOOST_CHECK_EQUAL(expected, workbook.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
