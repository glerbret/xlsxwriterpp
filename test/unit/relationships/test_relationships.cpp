/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/relationships.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(relationships)

BOOST_AUTO_TEST_CASE(assemble_xml_file_1)
{
  const std::string expected =
      // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
      "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
      "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
      "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
      "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.xml\"/>"
    "</Relationships>";
  // clang-format on

  xwpp::relationships_t relationships;
  relationships.add_document("/worksheet", "worksheets/sheet1.xml");
  relationships.add_document("/theme", "theme/theme1.xml");
  relationships.add_document("/styles", "styles.xml");
  relationships.add_document("/sharedStrings", "sharedStrings.xml");
  relationships.add_document("/calcChain", "calcChain.xml");

  BOOST_CHECK_EQUAL(expected, relationships.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(assemble_xml_file_2)
{
  const std::string expected =
      // clang-format off
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"www.foo.com\" TargetMode=\"External\"/>"
    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"link00.xlsx\" TargetMode=\"External\"/>"
  "</Relationships>";
  // clang-format on

  xwpp::relationships_t relationships;
  relationships.add_worksheet("/hyperlink", "www.foo.com", "External");
  relationships.add_worksheet("/hyperlink", "link00.xlsx", "External");

  BOOST_CHECK_EQUAL(expected, relationships.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
