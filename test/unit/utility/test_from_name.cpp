/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

#include <tuple>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(col_from_name)
{
  BOOST_CHECK_EQUAL(0, xwpp::col_from_name("A"));
  BOOST_CHECK_EQUAL(5, xwpp::col_from_name("F"));
}

BOOST_AUTO_TEST_CASE(cols_from_name)
{
  {
    const auto [first_col, last_col] = xwpp::cols_from_name("A:A");
    BOOST_CHECK_EQUAL(0, first_col);
    BOOST_CHECK_EQUAL(0, last_col);
  }

  {
    const auto [first_col, last_col] = xwpp::cols_from_name("A:E");
    BOOST_CHECK_EQUAL(0, first_col);
    BOOST_CHECK_EQUAL(4, last_col);
  }

  {
    const auto [first_col, last_col] = xwpp::cols_from_name("B:E");
    BOOST_CHECK_EQUAL(1, first_col);
    BOOST_CHECK_EQUAL(4, last_col);
  }

  {
    const auto [first_col, last_col] = xwpp::cols_from_name("A");
    BOOST_CHECK_EQUAL(0, first_col);
    BOOST_CHECK_EQUAL(0, last_col);
  }

  {
    const auto [first_col, last_col] = xwpp::cols_from_name("E");
    BOOST_CHECK_EQUAL(4, first_col);
    BOOST_CHECK_EQUAL(4, last_col);
  }
}

BOOST_AUTO_TEST_CASE(rows_from_name)
{
  {
    const auto [first_row, last_row] = xwpp::rows_from_name("1:1");
    BOOST_CHECK_EQUAL(0, first_row);
    BOOST_CHECK_EQUAL(0, last_row);
  }

  {
    const auto [first_row, last_row] = xwpp::rows_from_name("1:5");
    BOOST_CHECK_EQUAL(0, first_row);
    BOOST_CHECK_EQUAL(4, last_row);
  }

  {
    const auto [first_row, last_row] = xwpp::rows_from_name("5:1");
    BOOST_CHECK_EQUAL(4, first_row);
    BOOST_CHECK_EQUAL(0, last_row);
  }

  {
    const auto [first_row, last_row] = xwpp::rows_from_name("5");
    BOOST_CHECK_EQUAL(4, first_row);
    BOOST_CHECK_EQUAL(4, last_row);
  }
}

BOOST_AUTO_TEST_CASE(cell_from_name)
{
  {
    const auto [row_num, col_num] = xwpp::cell_from_name("A1");
    BOOST_CHECK_EQUAL(0, row_num);
    BOOST_CHECK_EQUAL(0, col_num);
  }

  {
    const auto [row_num, col_num] = xwpp::cell_from_name("A5");
    BOOST_CHECK_EQUAL(4, row_num);
    BOOST_CHECK_EQUAL(0, col_num);
  }

  {
    const auto [row_num, col_num] = xwpp::cell_from_name("F1");
    BOOST_CHECK_EQUAL(0, row_num);
    BOOST_CHECK_EQUAL(5, col_num);
  }

  {
    const auto [row_num, col_num] = xwpp::cell_from_name("F5");
    BOOST_CHECK_EQUAL(4, row_num);
    BOOST_CHECK_EQUAL(5, col_num);
  }

  {
    const auto [row_num, col_num] = xwpp::cell_from_name("AA12");
    BOOST_CHECK_EQUAL(11, row_num);
    BOOST_CHECK_EQUAL(26, col_num);
  }
}

BOOST_AUTO_TEST_CASE(range_from_name)
{
  {
    const auto [first_row, first_col, last_row, last_col] = xwpp::range_from_name("A1:B2");
    BOOST_CHECK_EQUAL(0, first_row);
    BOOST_CHECK_EQUAL(0, first_col);
    BOOST_CHECK_EQUAL(1, last_row);
    BOOST_CHECK_EQUAL(1, last_col);
  }

  {
    const auto [first_row, first_col, last_row, last_col] = xwpp::range_from_name("B2:AA50");
    BOOST_CHECK_EQUAL(1, first_row);
    BOOST_CHECK_EQUAL(1, first_col);
    BOOST_CHECK_EQUAL(49, last_row);
    BOOST_CHECK_EQUAL(26, last_col);
  }

  {
    const auto [first_row, first_col, last_row, last_col] = xwpp::range_from_name("A1");
    BOOST_CHECK_EQUAL(0, first_row);
    BOOST_CHECK_EQUAL(0, first_col);
    BOOST_CHECK_EQUAL(0, last_row);
    BOOST_CHECK_EQUAL(0, last_col);
  }

  {
    const auto [first_row, first_col, last_row, last_col] = xwpp::range_from_name("B2");
    BOOST_CHECK_EQUAL(1, first_row);
    BOOST_CHECK_EQUAL(1, first_col);
    BOOST_CHECK_EQUAL(1, last_row);
    BOOST_CHECK_EQUAL(1, last_col);
  }
}

BOOST_AUTO_TEST_SUITE_END()
