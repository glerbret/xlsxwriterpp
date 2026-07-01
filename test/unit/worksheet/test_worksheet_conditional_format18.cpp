/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"
#include "xwpp/workbook.h"
#include "xwpp/worksheet.h"

#include <string>

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(worksheet)

BOOST_AUTO_TEST_CASE(condtional_format18)
{
  const std::string expected =
      // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1:A12\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1\" spans=\"1:1\">"
          "<c r=\"A1\">"
            "<v>1</v>"
          "</c>"
        "</row>"
        "<row r=\"2\" spans=\"1:1\">"
          "<c r=\"A2\">"
            "<v>2</v>"
          "</c>"
        "</row>"
        "<row r=\"3\" spans=\"1:1\">"
          "<c r=\"A3\">"
            "<v>3</v>"
          "</c>"
        "</row>"
        "<row r=\"4\" spans=\"1:1\">"
          "<c r=\"A4\">"
            "<v>4</v>"
          "</c>"
        "</row>"
        "<row r=\"5\" spans=\"1:1\">"
          "<c r=\"A5\">"
            "<v>5</v>"
          "</c>"
        "</row>"
        "<row r=\"6\" spans=\"1:1\">"
          "<c r=\"A6\">"
            "<v>6</v>"
          "</c>"
        "</row>"
        "<row r=\"7\" spans=\"1:1\">"
          "<c r=\"A7\">"
            "<v>7</v>"
          "</c>"
        "</row>"
        "<row r=\"8\" spans=\"1:1\">"
          "<c r=\"A8\">"
            "<v>8</v>"
          "</c>"
        "</row>"
        "<row r=\"9\" spans=\"1:1\">"
          "<c r=\"A9\">"
            "<v>9</v>"
          "</c>"
        "</row>"
        "<row r=\"10\" spans=\"1:1\">"
          "<c r=\"A10\">"
            "<v>10</v>"
          "</c>"
        "</row>"
        "<row r=\"11\" spans=\"1:1\">"
          "<c r=\"A11\">"
            "<v>11</v>"
          "</c>"
        "</row>"
        "<row r=\"12\" spans=\"1:1\">"
          "<c r=\"A12\">"
            "<v>12</v>"
          "</c>"
        "</row>"
      "</sheetData>"
      "<conditionalFormatting sqref=\"A3:A4 A1 A6:A8 A10 A12\">"
        "<cfRule type=\"colorScale\" priority=\"1\">"
          "<colorScale>"
            "<cfvo type=\"min\" val=\"0\"/>"
            "<cfvo type=\"percentile\" val=\"50\"/>"
            "<cfvo type=\"max\" val=\"0\"/>"
            "<color rgb=\"FFF8696B\"/>"
            "<color rgb=\"FFFFEB84\"/>"
            "<color rgb=\"FF63BE7B\"/>"
          "</colorScale>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();

  worksheet.write_number(CELL("A1"), 1);
  worksheet.write_number(CELL("A2"), 2);
  worksheet.write_number(CELL("A3"), 3);
  worksheet.write_number(CELL("A4"), 4);
  worksheet.write_number(CELL("A5"), 5);
  worksheet.write_number(CELL("A6"), 6);
  worksheet.write_number(CELL("A7"), 7);
  worksheet.write_number(CELL("A8"), 8);
  worksheet.write_number(CELL("A9"), 9);
  worksheet.write_number(CELL("A10"), 10);
  worksheet.write_number(CELL("A11"), 11);
  worksheet.write_number(CELL("A12"), 12);

  xwpp::conditional_format_t conditional_format{
      .type_        = xwpp::conditional_format_types_t::THREE_COLOR_SCALE,
      .multi_range_ = "A3:A4 A1 A6:A8 A10 A12",
  };
  worksheet.conditional_format_range(RANGE("A1:A12"), conditional_format);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
