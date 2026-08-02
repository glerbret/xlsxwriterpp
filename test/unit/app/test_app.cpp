/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/app.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(app)

BOOST_AUTO_TEST_CASE(assemble_xml_file_1)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
        "<Application>Microsoft Excel</Application>"
        "<DocSecurity>0</DocSecurity>"
        "<ScaleCrop>false</ScaleCrop>"
        "<HeadingPairs>"
          "<vt:vector size=\"2\" baseType=\"variant\">"
            "<vt:variant>"
              "<vt:lpstr>Worksheets</vt:lpstr>"
            "</vt:variant>"
            "<vt:variant>"
              "<vt:i4>1</vt:i4>"
            "</vt:variant>"
          "</vt:vector>"
        "</HeadingPairs>"
        "<TitlesOfParts>"
          "<vt:vector size=\"1\" baseType=\"lpstr\">"
            "<vt:lpstr>Sheet1</vt:lpstr>"
          "</vt:vector>"
        "</TitlesOfParts>"
        "<Company>"
        "</Company>"
        "<LinksUpToDate>false</LinksUpToDate>"
        "<SharedDoc>false</SharedDoc>"
        "<HyperlinksChanged>false</HyperlinksChanged>"
        "<AppVersion>12.0000</AppVersion>"
      "</Properties>";
  // clang-format on

  xwpp::app_t app;
  app.add_part_name("Sheet1");
  app.add_heading_pair("Worksheets", "1");

  BOOST_CHECK_EQUAL(expected, app.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(assemble_xml_file_2)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
        "<Application>Microsoft Excel</Application>"
        "<DocSecurity>0</DocSecurity>"
        "<ScaleCrop>false</ScaleCrop>"
        "<HeadingPairs>"
          "<vt:vector size=\"2\" baseType=\"variant\">"
            "<vt:variant>"
              "<vt:lpstr>Worksheets</vt:lpstr>"
            "</vt:variant>"
            "<vt:variant>"
              "<vt:i4>2</vt:i4>"
            "</vt:variant>"
          "</vt:vector>"
        "</HeadingPairs>"
        "<TitlesOfParts>"
          "<vt:vector size=\"2\" baseType=\"lpstr\">"
            "<vt:lpstr>Sheet1</vt:lpstr>"
            "<vt:lpstr>Sheet2</vt:lpstr>"
          "</vt:vector>"
        "</TitlesOfParts>"
        "<Company>"
        "</Company>"
        "<LinksUpToDate>false</LinksUpToDate>"
        "<SharedDoc>false</SharedDoc>"
        "<HyperlinksChanged>false</HyperlinksChanged>"
        "<AppVersion>12.0000</AppVersion>"
      "</Properties>";
  // clang-format on

  xwpp::app_t app;
  app.add_part_name("Sheet1");
  app.add_part_name("Sheet2");
  app.add_heading_pair("Worksheets", "2");

  BOOST_CHECK_EQUAL(expected, app.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(assemble_xml_file_3)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
        "<Application>Microsoft Excel</Application>"
        "<DocSecurity>0</DocSecurity>"
        "<ScaleCrop>false</ScaleCrop>"
        "<HeadingPairs>"
          "<vt:vector size=\"4\" baseType=\"variant\">"
            "<vt:variant>"
              "<vt:lpstr>Worksheets</vt:lpstr>"
            "</vt:variant>"
            "<vt:variant>"
              "<vt:i4>1</vt:i4>"
            "</vt:variant>"
            "<vt:variant>"
              "<vt:lpstr>Named Ranges</vt:lpstr>"
            "</vt:variant>"
            "<vt:variant>"
              "<vt:i4>1</vt:i4>"
            "</vt:variant>"
          "</vt:vector>"
        "</HeadingPairs>"
        "<TitlesOfParts>"
          "<vt:vector size=\"2\" baseType=\"lpstr\">"
            "<vt:lpstr>Sheet1</vt:lpstr>"
            "<vt:lpstr>Sheet1!Print_Titles</vt:lpstr>"
          "</vt:vector>"
        "</TitlesOfParts>"
        "<Company>"
        "</Company>"
        "<LinksUpToDate>false</LinksUpToDate>"
        "<SharedDoc>false</SharedDoc>"
        "<HyperlinksChanged>false</HyperlinksChanged>"
        "<AppVersion>12.0000</AppVersion>"
      "</Properties>";
  // clang-format on

  xwpp::app_t app;
  app.add_part_name("Sheet1");
  app.add_part_name("Sheet1!Print_Titles");
  app.add_heading_pair("Worksheets", "1");
  app.add_heading_pair("Named Ranges", "1");

  BOOST_CHECK_EQUAL(expected, app.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
