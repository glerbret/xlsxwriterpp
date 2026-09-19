/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(reorder_index_col)
{
  {
    xwpp::col_num_t first_col = 0;
    xwpp::col_num_t last_col = 5;

    xwpp::reorder_index(first_col, last_col);

    BOOST_CHECK_EQUAL(xwpp::col_num_t{0}, first_col);
    BOOST_CHECK_EQUAL(xwpp::col_num_t{5}, last_col);
  }

  {
    xwpp::col_num_t first_col = 5;
    xwpp::col_num_t last_col = 0;

    xwpp::reorder_index(first_col, last_col);

    BOOST_CHECK_EQUAL(xwpp::col_num_t{0}, first_col);
    BOOST_CHECK_EQUAL(xwpp::col_num_t{5}, last_col);
  }
}

BOOST_AUTO_TEST_CASE(reorder_index_row)
{
  {
    xwpp::row_num_t first_row = 0;
    xwpp::row_num_t last_row = 5;

    xwpp::reorder_index(first_row, last_row);

    BOOST_CHECK_EQUAL(xwpp::row_num_t{0}, first_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{5}, last_row);
  }

  {
    xwpp::row_num_t first_row = 5;
    xwpp::row_num_t last_row = 0;

    xwpp::reorder_index(first_row, last_row);

    BOOST_CHECK_EQUAL(xwpp::row_num_t{0}, first_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{5}, last_row);
  }
}

BOOST_AUTO_TEST_CASE(reorder_index_row_col)
{
  {
    xwpp::row_num_t first_row = 0;
    xwpp::row_num_t last_row = 5;
    xwpp::col_num_t first_col = 1;
    xwpp::col_num_t last_col = 4;

    xwpp::reorder_index(first_row, last_row, first_col, last_col);

    BOOST_CHECK_EQUAL(xwpp::row_num_t{0}, first_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{5}, last_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{1}, first_col);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{4}, last_col);
  }

  {
    xwpp::row_num_t first_row = 5;
    xwpp::row_num_t last_row = 0;
    xwpp::col_num_t first_col = 1;
    xwpp::col_num_t last_col = 4;

    xwpp::reorder_index(first_row, last_row, first_col, last_col);

    BOOST_CHECK_EQUAL(xwpp::row_num_t{0}, first_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{5}, last_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{1}, first_col);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{4}, last_col);
  }

  {
    xwpp::row_num_t first_row = 0;
    xwpp::row_num_t last_row = 5;
    xwpp::col_num_t first_col = 4;
    xwpp::col_num_t last_col = 1;

    xwpp::reorder_index(first_row, last_row, first_col, last_col);

    BOOST_CHECK_EQUAL(xwpp::row_num_t{0}, first_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{5}, last_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{1}, first_col);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{4}, last_col);
  }

  {
    xwpp::row_num_t first_row = 5;
    xwpp::row_num_t last_row = 0;
    xwpp::col_num_t first_col = 4;
    xwpp::col_num_t last_col = 1;

    xwpp::reorder_index(first_row, last_row, first_col, last_col);

    BOOST_CHECK_EQUAL(xwpp::row_num_t{0}, first_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{5}, last_row);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{1}, first_col);
    BOOST_CHECK_EQUAL(xwpp::row_num_t{4}, last_col);
  }
}

BOOST_AUTO_TEST_SUITE_END()
