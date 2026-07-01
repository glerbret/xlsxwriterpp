/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/drawing.h"
#include "xwpp/worksheet.h"

#include <string>

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(drawing)

BOOST_AUTO_TEST_CASE(assemble_xml_file)
{
  const std::string expected =
      // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
      "<xdr:twoCellAnchor editAs=\"oneCell\">"
        "<xdr:from>"
          "<xdr:col>2</xdr:col>"
          "<xdr:colOff>0</xdr:colOff>"
          "<xdr:row>1</xdr:row>"
          "<xdr:rowOff>0</xdr:rowOff>"
        "</xdr:from>"
        "<xdr:to>"
          "<xdr:col>3</xdr:col>"
          "<xdr:colOff>533257</xdr:colOff>"
          "<xdr:row>6</xdr:row>"
          "<xdr:rowOff>190357</xdr:rowOff>"
        "</xdr:to>"
        "<xdr:pic>"
          "<xdr:nvPicPr>"
            "<xdr:cNvPr id=\"2\" name=\"Picture 1\" descr=\"republic.png\"/>"
            "<xdr:cNvPicPr>"
              "<a:picLocks noChangeAspect=\"1\"/>"
            "</xdr:cNvPicPr>"
          "</xdr:nvPicPr>"
          "<xdr:blipFill>"
            "<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId1\"/>"
            "<a:stretch>"
              "<a:fillRect/>"
            "</a:stretch>"
          "</xdr:blipFill>"
          "<xdr:spPr>"
            "<a:xfrm>"
              "<a:off x=\"1219200\" y=\"190500\"/>"
              "<a:ext cx=\"1142857\" cy=\"1142857\"/>"
            "</a:xfrm>"
            "<a:prstGeom prst=\"rect\">"
              "<a:avLst/>"
            "</a:prstGeom>"
          "</xdr:spPr>"
        "</xdr:pic>"
        "<xdr:clientData/>"
      "</xdr:twoCellAnchor>"
    "</xdr:wsDr>";
  // clang-format on

  xwpp::drawing_t drawing;
  const xwpp::drawing_object_t drawing_object{
      .type_         = xwpp::drawing_types_t::IMAGE,
      .anchor_       = static_cast<uint8_t>(xwpp::object_position_t::MOVE_DONT_SIZE),
      .from_         = {.col_ = 2, .row_ = 1, .col_offset_ = 0,      .row_offset_ = 0     },
      .to_           = {.col_ = 3, .row_ = 6, .col_offset_ = 533257, .row_offset_ = 190357},
      .col_absolute_ = 1219200,
      .row_absolute_ = 190500,
      .width_        = 1142857,
      .height_       = 1142857,
      .rel_index_    = 1,
      .description_  = "republic.png",
  };
  drawing.add_drawing_object(drawing_object);

  BOOST_CHECK_EQUAL(expected, drawing.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
