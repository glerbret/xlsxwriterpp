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

BOOST_AUTO_TEST_CASE(cols_from_name)
{
  {
    const auto[first_col, last_col] = xwpp::cols_from_name("A:A");
    BOOST_CHECK_EQUAL(0, first_col);
    BOOST_CHECK_EQUAL(0, last_col);
  }

  {
    const auto[first_col, last_col] = xwpp::cols_from_name("A:E");
    BOOST_CHECK_EQUAL(0, first_col);
    BOOST_CHECK_EQUAL(4, last_col);
  }

  {
    const auto[first_col, last_col] = xwpp::cols_from_name("A");
    BOOST_CHECK_EQUAL(0, first_col);
    BOOST_CHECK_EQUAL(0, last_col);
  }

  {
    const auto[first_col, last_col] = xwpp::cols_from_name("E");
    BOOST_CHECK_EQUAL(4, first_col);
    BOOST_CHECK_EQUAL(4, last_col);
  }
}

BOOST_AUTO_TEST_SUITE_END()
