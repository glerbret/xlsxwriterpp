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

BOOST_AUTO_TEST_CASE(data_bar11)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" mc:Ignorable=\"x14ac\">"
      "<dimension ref=\"A1\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>"
      "<sheetData/>"
      "<conditionalFormatting sqref=\"A1\">"
        "<cfRule type=\"dataBar\" priority=\"1\">"
          "<dataBar>"
            "<cfvo type=\"formula\" val=\"$B$1\"/>"
            "<cfvo type=\"max\"/>"
            "<color rgb=\"FF638EC6\"/>"
          "</dataBar>"
          "<extLst>"
            "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\">"
              "<x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>"
            "</ext>"
          "</extLst>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<conditionalFormatting sqref=\"A2:B2\">"
        "<cfRule type=\"dataBar\" priority=\"2\">"
          "<dataBar>"
            "<cfvo type=\"formula\" val=\"$B$1\"/>"
            "<cfvo type=\"formula\" val=\"$C$1\"/>"
            "<color rgb=\"FF63C384\"/>"
          "</dataBar>"
          "<extLst>"
            "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\">"
              "<x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>"
            "</ext>"
          "</extLst>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<conditionalFormatting sqref=\"A3:C3\">"
        "<cfRule type=\"dataBar\" priority=\"3\">"
          "<dataBar>"
            "<cfvo type=\"percentile\" val=\"10\"/>"
            "<cfvo type=\"percentile\" val=\"90\"/>"
            "<color rgb=\"FFFF555A\"/>"
          "</dataBar>"
          "<extLst>"
            "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\">"
              "<x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id>"
            "</ext>"
          "</extLst>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "<extLst>"
        "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\">"
          "<x14:conditionalFormattings>"
            "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
              "<x14:cfRule type=\"dataBar\" id=\"{DA7ABA51-AAAA-BBBB-0001-000000000001}\">"
                "<x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\" negativeBarBorderColorSameAsPositive=\"0\">"
                  "<x14:cfvo type=\"formula\">"
                    "<xm:f>$B$1</xm:f>"
                  "</x14:cfvo>"
                  "<x14:cfvo type=\"autoMax\"/>"
                  "<x14:borderColor rgb=\"FF638EC6\"/>"
                  "<x14:negativeFillColor rgb=\"FFFF0000\"/>"
                  "<x14:negativeBorderColor rgb=\"FFFF0000\"/>"
                  "<x14:axisColor rgb=\"FF000000\"/>"
                "</x14:dataBar>"
              "</x14:cfRule>"
              "<xm:sqref>A1</xm:sqref>"
            "</x14:conditionalFormatting>"
            "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
              "<x14:cfRule type=\"dataBar\" id=\"{DA7ABA51-AAAA-BBBB-0001-000000000002}\">"
                "<x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\" negativeBarBorderColorSameAsPositive=\"0\">"
                  "<x14:cfvo type=\"formula\">"
                    "<xm:f>$B$1</xm:f>"
                  "</x14:cfvo>"
                  "<x14:cfvo type=\"formula\">"
                    "<xm:f>$C$1</xm:f>"
                  "</x14:cfvo>"
                  "<x14:borderColor rgb=\"FF63C384\"/>"
                  "<x14:negativeFillColor rgb=\"FFFF0000\"/>"
                  "<x14:negativeBorderColor rgb=\"FFFF0000\"/>"
                  "<x14:axisColor rgb=\"FF000000\"/>"
                "</x14:dataBar>"
              "</x14:cfRule>"
              "<xm:sqref>A2:B2</xm:sqref>"
            "</x14:conditionalFormatting>"
            "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
              "<x14:cfRule type=\"dataBar\" id=\"{DA7ABA51-AAAA-BBBB-0001-000000000003}\">"
                "<x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\" negativeBarBorderColorSameAsPositive=\"0\">"
                  "<x14:cfvo type=\"percentile\">"
                    "<xm:f>10</xm:f>"
                  "</x14:cfvo>"
                  "<x14:cfvo type=\"percentile\">"
                    "<xm:f>90</xm:f>"
                  "</x14:cfvo>"
                  "<x14:borderColor rgb=\"FFFF555A\"/>"
                  "<x14:negativeFillColor rgb=\"FFFF0000\"/>"
                  "<x14:negativeBorderColor rgb=\"FFFF0000\"/>"
                  "<x14:axisColor rgb=\"FF000000\"/>"
                "</x14:dataBar>"
              "</x14:cfRule>"
              "<xm:sqref>A3:C3</xm:sqref>"
            "</x14:conditionalFormatting>"
          "</x14:conditionalFormattings>"
        "</ext>"
      "</extLst>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();

  xwpp::conditional_format_t conditional_format{
    .type_             = xwpp::conditional_format_types_t::DATA_BAR,
    .min_value_string_ = "=$B$1",
    .min_rule_type_    = xwpp::conditional_format_rule_types_t::FORMULA,
    .data_bar_2010_    = true,
  };
  worksheet.conditional_format_cell(CELL("A1"), conditional_format);

  conditional_format = xwpp::conditional_format_t{
    .type_             = xwpp::conditional_format_types_t::DATA_BAR,
    .min_value_string_ = "=$B$1",
    .min_rule_type_    = xwpp::conditional_format_rule_types_t::FORMULA,
    .max_value_string_ = "=$C$1",
    .max_rule_type_    = xwpp::conditional_format_rule_types_t::FORMULA,
    .bar_color_        = xwpp::color_t(0x63C384),
    .data_bar_2010_    = true,
  };
  worksheet.conditional_format_range(RANGE("A2:B2"), conditional_format);

  conditional_format = xwpp::conditional_format_t{
    .type_          = xwpp::conditional_format_types_t::DATA_BAR,
    .min_value_     = 10,
    .min_rule_type_ = xwpp::conditional_format_rule_types_t::PERCENTILE,
    .max_value_     = 90,
    .max_rule_type_ = xwpp::conditional_format_rule_types_t::PERCENTILE,
    .bar_color_     = xwpp::color_t(0xFF555A),
    .data_bar_2010_ = true,
  };
  worksheet.conditional_format_range(RANGE("A3:C3"), conditional_format);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

// percentile" val="90"/><color rgb="FFFF555A"/></dataBar><extLst><ext
// xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
// uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id></ext></extLst></cfRule></conditionalFormatting><pageMargins
// left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><extLst><ext
// xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
// uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting
// xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar"
// id="{DA7ABA51-AAAA-BBBB-0001-000000000001}"><x14:dataBar minLength="0" maxLength="100" border="1"
// negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="formula"><xm:f>$B$1</xm:f></x14:cfvo><x14:cfvo
// type="autoMax"/><x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor
// rgb="FFFF0000"/><x14:axisColor
// rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A1</xm:sqref></x14:conditionalFormatting><x14:conditionalFormatting
// xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar"
// id="{DA7ABA51-AAAA-BBBB-0001-000000000002}"><x14:dataBar minLength="0" maxLength="100" border="1"
// negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="formula"><xm:f>$B$1</xm:f></x14:cfvo><x14:cfvo
// type="formula"><xm:f>$C$1</xm:f></x14:cfvo><x14:borderColor rgb="FF63C384"/><x14:negativeFillColor
// rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor
// rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A2:B2</xm:sqref></x14:conditionalFormatting><x14:conditionalFormatting
// xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar"
// id="{DA7ABA51-AAAA-BBBB-0001-000000000003}"><x14:dataBar minLength="0" maxLength="100" border="1"
// negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="percentile"><xm:f>10</xm:f></x14:cfvo><x14:cfvo
// type="percentile"><xm:f>90</xm:f></x14:cfvo><x14:borderColor rgb="FFFF555A"/><x14:negativeFillColor
// rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor
// rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A3:C3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>
// percentile" val="99"/><color rgb="FFFF555A"/></dataBar><extLst><ext
// xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
// uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id></ext></extLst></cfRule></conditionalFormatting><pageMargins
// left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><extLst><ext
// xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
// uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting
// xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar"
// id="{DA7ABA51-AAAA-BBBB-0001-000000000001}"><x14:dataBar minLength="0" maxLength="100" border="1"
// negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="formula"><xm:f>$B$1</xm:f></x14:cfvo><x14:cfvo
// type="autoMax"/><x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor
// rgb="FFFF0000"/><x14:axisColor
// rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A1</xm:sqref></x14:conditionalFormatting><x14:conditionalFormatting
// xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar"
// id="{DA7ABA51-AAAA-BBBB-0001-000000000002}"><x14:dataBar minLength="0" maxLength="100" border="1"
// negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="formula"><xm:f>$B$1</xm:f></x14:cfvo><x14:cfvo
// type="formula"><xm:f>$C$1</xm:f></x14:cfvo><x14:borderColor rgb="FF63C384"/><x14:negativeFillColor
// rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor
// rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A2:B2</xm:sqref></x14:conditionalFormatting><x14:conditionalFormatting
// xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar"
// id="{DA7ABA51-AAAA-BBBB-0001-000000000003}"><x14:dataBar minLength="0" maxLength="100" border="1"
// negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="percentile"><xm:f>10</xm:f></x14:cfvo><x14:cfvo
// type="percentile"><xm:f>99</xm:f></x14:cfvo><x14:borderColor rgb="FFFF555A"/><x14:negativeFillColor
// rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor
// rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A3:C3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>

BOOST_AUTO_TEST_SUITE_END()
