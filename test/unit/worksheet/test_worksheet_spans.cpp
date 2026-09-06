/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/exception.h"
#include "xwpp/utility.h"
#include "xwpp/workbook.h"
#include "xwpp/worksheet.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(worksheet)

BOOST_AUTO_TEST_CASE(spans01)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"B3\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"3\" spans=\"2:2\">"
          "<c r=\"B3\">"
            "<v>2000</v>"
          "</c>"
        "</row>"
      "</sheetData>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();

  worksheet.write_number(2, 1, 2000);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(spans02)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1048576\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1048576\" spans=\"1:1\">"
          "<c r=\"A1048576\">"
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

  worksheet.write_number(1048575, 0, 123);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(spans03)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"XFD1\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1\" spans=\"16384:16384\">"
          "<c r=\"XFD1\">"
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

  worksheet.write_number(0, 16383, 123);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(spans04)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"XFD1048576\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1048576\" spans=\"16384:16384\">"
          "<c r=\"XFD1048576\">"
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

  worksheet.write_number(1048575, 16383, 123);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(spans05)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1:T20\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1\" spans=\"1:16\">"
          "<c r=\"A1\">"
            "<v>1</v>"
          "</c>"
        "</row>"
        "<row r=\"2\" spans=\"1:16\">"
          "<c r=\"B2\">"
            "<v>2</v>"
          "</c>"
        "</row>"
        "<row r=\"3\" spans=\"1:16\">"
          "<c r=\"C3\">"
            "<v>3</v>"
          "</c>"
        "</row>"
        "<row r=\"4\" spans=\"1:16\">"
          "<c r=\"D4\">"
            "<v>4</v>"
          "</c>"
        "</row>"
        "<row r=\"5\" spans=\"1:16\">"
          "<c r=\"E5\">"
            "<v>5</v>"
          "</c>"
        "</row>"
        "<row r=\"6\" spans=\"1:16\">"
          "<c r=\"F6\">"
            "<v>6</v>"
          "</c>"
        "</row>"
        "<row r=\"7\" spans=\"1:16\">"
          "<c r=\"G7\">"
            "<v>7</v>"
          "</c>"
        "</row>"
        "<row r=\"8\" spans=\"1:16\">"
          "<c r=\"H8\">"
            "<v>8</v>"
          "</c>"
        "</row>"
        "<row r=\"9\" spans=\"1:16\">"
          "<c r=\"I9\">"
            "<v>9</v>"
          "</c>"
        "</row>"
        "<row r=\"10\" spans=\"1:16\">"
          "<c r=\"J10\">"
            "<v>10</v>"
          "</c>"
        "</row>"
        "<row r=\"11\" spans=\"1:16\">"
          "<c r=\"K11\">"
            "<v>11</v>"
          "</c>"
        "</row>"
        "<row r=\"12\" spans=\"1:16\">"
          "<c r=\"L12\">"
            "<v>12</v>"
          "</c>"
        "</row>"
        "<row r=\"13\" spans=\"1:16\">"
          "<c r=\"M13\">"
            "<v>13</v>"
          "</c>"
        "</row>"
        "<row r=\"14\" spans=\"1:16\">"
          "<c r=\"N14\">"
            "<v>14</v>"
          "</c>"
        "</row>"
        "<row r=\"15\" spans=\"1:16\">"
          "<c r=\"O15\">"
            "<v>15</v>"
          "</c>"
        "</row>"
        "<row r=\"16\" spans=\"1:16\">"
          "<c r=\"P16\">"
            "<v>16</v>"
          "</c>"
        "</row>"
        "<row r=\"17\" spans=\"17:20\">"
          "<c r=\"Q17\">"
            "<v>17</v>"
          "</c>"
        "</row>"
        "<row r=\"18\" spans=\"17:20\">"
          "<c r=\"R18\">"
            "<v>18</v>"
          "</c>"
        "</row>"
        "<row r=\"19\" spans=\"17:20\">"
          "<c r=\"S19\">"
            "<v>19</v>"
          "</c>"
        "</row>"
        "<row r=\"20\" spans=\"17:20\">"
          "<c r=\"T20\">"
            "<v>20</v>"
          "</c>"
        "</row>"
      "</sheetData>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();

  for(size_t i{0}; i < 20; i++)
  {
    worksheet.write_number(static_cast<xwpp::row_num_t>(i), static_cast<xwpp::col_num_t>(i),
                           static_cast<double>(i + 1));
  }

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(spans06)
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();

  BOOST_CHECK_THROW(worksheet.write_number(0, 16384, 123), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_number(1048576, 0, 123), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_number(1048576, 16384, 123), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_SUITE_END()
