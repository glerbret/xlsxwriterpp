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

BOOST_AUTO_TEST_CASE(condtional_format23)
{
  const std::string expected =
      // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1:A8\"/>"
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
      "</sheetData>"
      "<conditionalFormatting sqref=\"A1\">"
        "<cfRule type=\"iconSet\" priority=\"1\">"
          "<iconSet iconSet=\"3ArrowsGray\">"
            "<cfvo type=\"percent\" val=\"0\"/>"
            "<cfvo type=\"percent\" val=\"33\"/>"
            "<cfvo type=\"percent\" val=\"67\"/>"
          "</iconSet>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<conditionalFormatting sqref=\"A2\">"
        "<cfRule type=\"iconSet\" priority=\"2\">"
          "<iconSet>"
            "<cfvo type=\"percent\" val=\"0\"/>"
            "<cfvo type=\"percent\" val=\"33\"/>"
            "<cfvo type=\"percent\" val=\"67\"/>"
          "</iconSet>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<conditionalFormatting sqref=\"A3\">"
        "<cfRule type=\"iconSet\" priority=\"3\">"
          "<iconSet iconSet=\"3Signs\">"
            "<cfvo type=\"percent\" val=\"0\"/>"
            "<cfvo type=\"percent\" val=\"33\"/>"
            "<cfvo type=\"percent\" val=\"67\"/>"
          "</iconSet>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<conditionalFormatting sqref=\"A4\">"
        "<cfRule type=\"iconSet\" priority=\"4\">"
          "<iconSet iconSet=\"3Symbols2\">"
            "<cfvo type=\"percent\" val=\"0\"/>"
            "<cfvo type=\"percent\" val=\"33\"/>"
            "<cfvo type=\"percent\" val=\"67\"/>"
          "</iconSet>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<conditionalFormatting sqref=\"A5\">"
        "<cfRule type=\"iconSet\" priority=\"5\">"
          "<iconSet iconSet=\"4ArrowsGray\">"
            "<cfvo type=\"percent\" val=\"0\"/>"
            "<cfvo type=\"percent\" val=\"25\"/>"
            "<cfvo type=\"percent\" val=\"50\"/>"
            "<cfvo type=\"percent\" val=\"75\"/>"
          "</iconSet>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<conditionalFormatting sqref=\"A6\">"
        "<cfRule type=\"iconSet\" priority=\"6\">"
          "<iconSet iconSet=\"4Rating\">"
            "<cfvo type=\"percent\" val=\"0\"/>"
            "<cfvo type=\"percent\" val=\"25\"/>"
            "<cfvo type=\"percent\" val=\"50\"/>"
            "<cfvo type=\"percent\" val=\"75\"/>"
          "</iconSet>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<conditionalFormatting sqref=\"A7\">"
        "<cfRule type=\"iconSet\" priority=\"7\">"
          "<iconSet iconSet=\"5Arrows\">"
            "<cfvo type=\"percent\" val=\"0\"/>"
            "<cfvo type=\"percent\" val=\"20\"/>"
            "<cfvo type=\"percent\" val=\"40\"/>"
            "<cfvo type=\"percent\" val=\"60\"/>"
            "<cfvo type=\"percent\" val=\"80\"/>"
          "</iconSet>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<conditionalFormatting sqref=\"A8\">"
        "<cfRule type=\"iconSet\" priority=\"8\">"
          "<iconSet iconSet=\"5Rating\">"
            "<cfvo type=\"percent\" val=\"0\"/>"
            "<cfvo type=\"percent\" val=\"20\"/>"
            "<cfvo type=\"percent\" val=\"40\"/>"
            "<cfvo type=\"percent\" val=\"60\"/>"
            "<cfvo type=\"percent\" val=\"80\"/>"
          "</iconSet>"
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

  xwpp::conditional_format_t conditional_format{
      .type_       = xwpp::conditional_format_types_t::ICON_SETS,
      .icon_style_ = xwpp::conditional_icon_types_t::THREE_ARROWS_GRAY,
  };
  worksheet.conditional_format_cell(CELL("A1"), conditional_format);

  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::THREE_TRAFFIC_LIGHTS_UNRIMMED;
  worksheet.conditional_format_cell(CELL("A2"), conditional_format);

  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::THREE_SIGNS;
  worksheet.conditional_format_cell(CELL("A3"), conditional_format);

  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::THREE_SYMBOLS_UNCIRCLED;
  worksheet.conditional_format_cell(CELL("A4"), conditional_format);

  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::FOUR_ARROWS_GRAY;
  worksheet.conditional_format_cell(CELL("A5"), conditional_format);

  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::FOUR_RATINGS;
  worksheet.conditional_format_cell(CELL("A6"), conditional_format);

  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::FIVE_ARROWS_COLORED;
  worksheet.conditional_format_cell(CELL("A7"), conditional_format);

  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::FIVE_RATINGS;
  worksheet.conditional_format_cell(CELL("A8"), conditional_format);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
