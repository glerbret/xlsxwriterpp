/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"
#include "xwpp/workbook.h"
#include "xwpp/worksheet.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(worksheet)

BOOST_AUTO_TEST_CASE(condtional_format20)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1:A4\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1\" spans=\"1:1\">"
          "<c r=\"A1\">"
            "<v>10</v>"
          "</c>"
        "</row>"
        "<row r=\"2\" spans=\"1:1\">"
          "<c r=\"A2\">"
            "<v>20</v>"
          "</c>"
        "</row>"
        "<row r=\"3\" spans=\"1:1\">"
          "<c r=\"A3\">"
            "<v>30</v>"
          "</c>"
        "</row>"
        "<row r=\"4\" spans=\"1:1\">"
          "<c r=\"A4\">"
            "<v>40</v>"
          "</c>"
        "</row>"
      "</sheetData>"
      "<conditionalFormatting sqref=\"A1:A4\">"
        "<cfRule type=\"beginsWith\" operator=\"beginsWith\" priority=\"1\" text=\"b\">"
          "<formula>LEFT(A1,1)=\"b\"</formula>"
        "</cfRule>"
        "<cfRule type=\"beginsWith\" operator=\"beginsWith\" priority=\"2\" text=\"bc\">"
          "<formula>LEFT(A1,2)=\"bc\"</formula>"
        "</cfRule>"
        "<cfRule type=\"endsWith\" operator=\"endsWith\" priority=\"3\" text=\"z\">"
          "<formula>RIGHT(A1,1)=\"z\"</formula>"
        "</cfRule>"
        "<cfRule type=\"endsWith\" operator=\"endsWith\" priority=\"4\" text=\"yz\">"
          "<formula>RIGHT(A1,2)=\"yz\"</formula>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();

  worksheet.write("A1", 10);
  worksheet.write("A2", 20);
  worksheet.write("A3", 30);
  worksheet.write("A4", 40);

  xwpp::conditional_format_t conditional_format{
    .type         = xwpp::conditional_format_types_t::TEXT,
    .criteria     = xwpp::conditional_criteria_t::TEXT_BEGINS_WITH,
    .value_string = "b",
  };
  worksheet.conditional_format_range("A1:A4", conditional_format);

  conditional_format.type         = xwpp::conditional_format_types_t::TEXT;
  conditional_format.criteria     = xwpp::conditional_criteria_t::TEXT_BEGINS_WITH;
  conditional_format.value_string = "bc";
  worksheet.conditional_format_range("A1:A4", conditional_format);

  conditional_format.type         = xwpp::conditional_format_types_t::TEXT;
  conditional_format.criteria     = xwpp::conditional_criteria_t::TEXT_ENDS_WITH;
  conditional_format.value_string = "z";
  worksheet.conditional_format_range("A1:A4", conditional_format);

  conditional_format.type         = xwpp::conditional_format_types_t::TEXT;
  conditional_format.criteria     = xwpp::conditional_criteria_t::TEXT_ENDS_WITH;
  conditional_format.value_string = "yz";
  worksheet.conditional_format_range("A1:A4", conditional_format);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
