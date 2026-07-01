/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/workbook.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(workbook)

BOOST_AUTO_TEST_CASE(get_worksheet_by_name_default_name)
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  BOOST_CHECK_EQUAL(&worksheet, workbook.get_worksheet_by_name("Sheet1"));
}

BOOST_AUTO_TEST_CASE(get_worksheet_by_name_provided_name)
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet("FOO");

  BOOST_CHECK_EQUAL(&worksheet, workbook.get_worksheet_by_name("FOO"));
}

BOOST_AUTO_TEST_CASE(get_worksheet_by_name_different_name)
{
  xwpp::workbook_t workbook;
  workbook.add_worksheet();

  BOOST_TEST(workbook.get_worksheet_by_name("FOO") == nullptr);
}

BOOST_AUTO_TEST_CASE(get_worksheet_by_name_no_sheetname)
{
  xwpp::workbook_t workbook;

  BOOST_TEST(workbook.get_worksheet_by_name("FOO") == nullptr);
}

BOOST_AUTO_TEST_CASE(get_worksheet_by_name_no_name)
{
  xwpp::workbook_t workbook;
  workbook.add_worksheet();

  BOOST_TEST(workbook.get_worksheet_by_name("") == nullptr);
}

BOOST_AUTO_TEST_SUITE_END()
