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

BOOST_AUTO_TEST_CASE(bound_checks)
{
  const xwpp::row_num_t MAX_ROW = 1048576;
  const xwpp::col_num_t MAX_COL = 16384;

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  const xwpp::format_t* format = workbook.add_format();
  worksheet.select();

  BOOST_CHECK_THROW(worksheet.write_number(0, MAX_COL, 123), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_number(MAX_ROW, 0, 123), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_number(MAX_ROW, MAX_COL, 123), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_string(MAX_ROW, 0, "Foo"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_string(0, MAX_COL, "Foo"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_string(MAX_ROW, MAX_COL, "Foo"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_number(MAX_ROW, 0, 123), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_number(0, MAX_COL, 123), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_number(MAX_ROW, MAX_COL, 123), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_blank(MAX_ROW, 0, format), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_blank(0, MAX_COL, format), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_blank(MAX_ROW, MAX_COL, format), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_formula(MAX_ROW, 0, "=A1"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_formula(0, MAX_COL, "=A1"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_formula(MAX_ROW, MAX_COL, "=A1"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_array_formula(0, 0, 0, MAX_COL, "=A1"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_array_formula(0, 0, MAX_ROW, 0, "=A1"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_array_formula(0, MAX_COL, 0, 0, "=A1"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_array_formula(MAX_ROW, 0, 0, 0, "=A1"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_array_formula(MAX_ROW, MAX_COL, MAX_ROW, MAX_COL, "=A1"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.merge_range(0, 0, 0, MAX_COL, "Foo", nullptr), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.merge_range(0, 0, MAX_ROW, 0, "Foo", nullptr), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.merge_range(0, MAX_COL, 0, 0, "Foo", nullptr), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.merge_range(MAX_ROW, 0, 0, 0, "Foo", nullptr), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.set_column(6, MAX_COL, 17), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.set_column(MAX_COL, 6, 17), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_array_formula(0, 0, 0, 0, "{"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_array_formula(0, 0, 0, 0, "}"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_array_formula(0, 0, 0, 0, "{}"), xwpp::xwpp_exception_t);
  BOOST_CHECK_THROW(worksheet.write_array_formula(0, 0, 0, 0, "{=}"), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_SUITE_END()
