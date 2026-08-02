/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/chartsheet.h"
#include "xwpp/workbook.h"

#include <string>

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(chartsheet)

BOOST_AUTO_TEST_CASE(write_sheet_protection_no_password)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr/>"
        "<sheetViews>"
          "<sheetView workbookViewId=\"0\"/>"
        "</sheetViews>"
        "<sheetProtection content=\"1\" objects=\"1\"/>"
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "</chartsheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
  chartsheet.protect();

  BOOST_CHECK_EQUAL(expected, chartsheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(write_sheet_protection_password)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr/>"
        "<sheetViews>"
          "<sheetView workbookViewId=\"0\"/>"
        "</sheetViews>"
        "<sheetProtection password=\"83AF\" content=\"1\" objects=\"1\"/>"
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "</chartsheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
  chartsheet.protect("password");

  BOOST_CHECK_EQUAL(expected, chartsheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(write_sheet_protection_no_object)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr/>"
        "<sheetViews>"
          "<sheetView workbookViewId=\"0\"/>"
        "</sheetViews>"
        "<sheetProtection content=\"1\"/>"
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "</chartsheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
  xwpp::protection_t options     = {.no_objects_ = true};
  chartsheet.protect(options);

  BOOST_CHECK_EQUAL(expected, chartsheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(write_sheet_protection_no_content)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr/>"
        "<sheetViews>"
          "<sheetView workbookViewId=\"0\"/>"
        "</sheetViews>"
        "<sheetProtection objects=\"1\"/>"
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "</chartsheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
  xwpp::protection_t options     = {.no_content_ = true};
  chartsheet.protect(options);

  BOOST_CHECK_EQUAL(expected, chartsheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(write_sheet_protection_no_object_no_content)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr/>"
        "<sheetViews>"
          "<sheetView workbookViewId=\"0\"/>"
        "</sheetViews>"
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "</chartsheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
  xwpp::protection_t options     = {.no_content_ = true, .no_objects_ = true};
  chartsheet.protect(options);

  BOOST_CHECK_EQUAL(expected, chartsheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(write_sheet_protection_no_object_no_content_password)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr/>"
        "<sheetViews>"
          "<sheetView workbookViewId=\"0\"/>"
        "</sheetViews>"
        "<sheetProtection password=\"83AF\"/>"
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "</chartsheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
  xwpp::protection_t options     = {.no_content_ = true, .no_objects_ = true};
  chartsheet.protect("password", options);

  BOOST_CHECK_EQUAL(expected, chartsheet.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(write_sheet_protection_other_options)
{
  const std::string expected =
    // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr/>"
        "<sheetViews>"
          "<sheetView workbookViewId=\"0\"/>"
        "</sheetViews>"
        "<sheetProtection password=\"83AF\" content=\"1\" objects=\"1\"/>"
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "</chartsheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
  xwpp::protection_t options     = {
        .no_select_locked_cells_   = true,
        .no_select_unlocked_cells_ = true,
        .format_cells_             = true,
        .format_columns_           = true,
        .format_rows_              = true,
        .insert_columns_           = true,
        .insert_rows_              = true,
        .insert_hyperlinks_        = true,
        .delete_columns_           = true,
        .delete_rows_              = true,
        .sort_                     = true,
        .autofilter_               = true,
        .pivot_tables_             = true,
        .scenarios_                = true,
        .objects_                  = true,
  };
  chartsheet.protect("password", options);

  BOOST_CHECK_EQUAL(expected, chartsheet.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
