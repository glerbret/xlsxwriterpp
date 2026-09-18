/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/workbook.h"

#include "xwpp/exception.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(workbook)

BOOST_AUTO_TEST_CASE(get_chartsheet_by_name_default_name)
{
  xwpp::workbook_t workbook;
  workbook.add_chartsheet();

  BOOST_CHECK(workbook.is_chartsheet_present("Chart1"));
  BOOST_CHECK_NO_THROW((void)workbook.get_chartsheet_by_name("Chart1"));
}

BOOST_AUTO_TEST_CASE(get_chartsheet_by_name_provided_name)
{
  xwpp::workbook_t workbook;
  workbook.add_chartsheet("FOO");

  BOOST_CHECK(workbook.is_chartsheet_present("FOO"));
  BOOST_CHECK_NO_THROW((void)workbook.get_chartsheet_by_name("FOO"));
}

BOOST_AUTO_TEST_CASE(get_chartsheet_by_name_different_name)
{
  xwpp::workbook_t workbook;
  workbook.add_chartsheet();

  BOOST_CHECK(!workbook.is_chartsheet_present("FOO"));
  BOOST_CHECK_THROW((void)workbook.get_chartsheet_by_name("FOO"), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(get_chartsheet_by_name_no_sheetname)
{
  const xwpp::workbook_t workbook;

  BOOST_CHECK(!workbook.is_chartsheet_present("FOO"));
  BOOST_CHECK_THROW((void)workbook.get_chartsheet_by_name("FOO"), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(get_chartsheet_by_name_no_name)
{
  xwpp::workbook_t workbook;
  workbook.add_chartsheet();

  BOOST_CHECK(!workbook.is_chartsheet_present(""));
  BOOST_CHECK_THROW((void)workbook.get_chartsheet_by_name(""), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_SUITE_END()
