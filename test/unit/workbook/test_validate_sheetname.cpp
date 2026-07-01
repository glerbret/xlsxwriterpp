/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/workbook.h"

#include "xwpp/exception.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(workbook)

BOOST_AUTO_TEST_CASE(validate_sheetname_valid)
{
  const std::string sheetname = "123456789_123456789_123456789_1";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_NO_THROW(workbook.validate_sheetname(sheetname));
}

BOOST_AUTO_TEST_CASE(validate_sheetname_too_long)
{
  const std::string sheetname = "123456789_123456789_123456789_12";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_sheetname_invalid_char)
{
  const std::string sheetname = "Sheet[1]";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_sheetname_unbalanced_single_quote_start)
{
  const std::string sheetname = "'Sheet1";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_sheetname_unbalanced_single_quote_end)
{
  const std::string sheetname = "Sheet1'";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_worksheet_already_exists)
{
  const std::string sheetname = "Sheet1";
  xwpp::workbook_t workbook;
  workbook.add_worksheet(sheetname);

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_sheetname_unbalanced_apostrophe_start)
{
  const std::string sheetname = "'Sheet1";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_sheetname_unbalanced_apostrophe_end)
{
  const std::string sheetname = "Sheet1'";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_worksheet_already_exists_case_insensitive_1)
{
  const std::string sheetname = "Sheet1";
  xwpp::workbook_t workbook;
  workbook.add_worksheet(sheetname);

  BOOST_REQUIRE_THROW(workbook.validate_sheetname("sheet1"), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_worksheet_already_exists_case_insensitive_2)
{
  const std::string sheetname = "Café";
  xwpp::workbook_t workbook;
  workbook.add_worksheet(sheetname);

  BOOST_REQUIRE_THROW(workbook.validate_sheetname("café"), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_worksheet_already_exists_case_insensitive_3)
{
  const std::string sheetname = "abcde";
  xwpp::workbook_t workbook;
  workbook.add_worksheet(sheetname);

  BOOST_REQUIRE_THROW(workbook.validate_sheetname("ABCDE"), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(validate_worksheet_empty_name)
{
  const std::string sheetname = "";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_SUITE_END()
