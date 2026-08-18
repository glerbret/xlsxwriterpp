/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(name_to_row)
{
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row(""));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row("A"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row("A0"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row("A1"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row("$A$1"));
  BOOST_CHECK_EQUAL(1, xwpp::name_to_row("B2"));
  BOOST_CHECK_EQUAL(2, xwpp::name_to_row("C3"));
  BOOST_CHECK_EQUAL(9, xwpp::name_to_row("J10"));
  BOOST_CHECK_EQUAL(24, xwpp::name_to_row("Y25"));
  BOOST_CHECK_EQUAL(25, xwpp::name_to_row("Z26"));
  BOOST_CHECK_EQUAL(26, xwpp::name_to_row("AA27"));
  BOOST_CHECK_EQUAL(254, xwpp::name_to_row("IU255"));
  BOOST_CHECK_EQUAL(255, xwpp::name_to_row("IV256"));
  BOOST_CHECK_EQUAL(256, xwpp::name_to_row("IW257"));
  BOOST_CHECK_EQUAL(16383, xwpp::name_to_row("XFD16384"));
  BOOST_CHECK_EQUAL(16384, xwpp::name_to_row("XFE16385"));
  BOOST_CHECK_EQUAL(1048576, xwpp::name_to_row("XFE1048577"));
  BOOST_CHECK_EQUAL(1048576, xwpp::name_to_row("$XFE$1048577"));
}

BOOST_AUTO_TEST_CASE(name_to_row_2)
{
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row_2(""));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row_2("A1:A"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row_2("A1:A0"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row_2("A1:A1"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_row_2("A1:$A$1"));
  BOOST_CHECK_EQUAL(1, xwpp::name_to_row_2("A1:B2"));
  BOOST_CHECK_EQUAL(2, xwpp::name_to_row_2("A1:C3"));
  BOOST_CHECK_EQUAL(9, xwpp::name_to_row_2("A1:J10"));
  BOOST_CHECK_EQUAL(24, xwpp::name_to_row_2("A1:Y25"));
  BOOST_CHECK_EQUAL(25, xwpp::name_to_row_2("A1:Z26"));
  BOOST_CHECK_EQUAL(26, xwpp::name_to_row_2("A1:AA27"));
  BOOST_CHECK_EQUAL(254, xwpp::name_to_row_2("A1:IU255"));
  BOOST_CHECK_EQUAL(255, xwpp::name_to_row_2("A1:IV256"));
  BOOST_CHECK_EQUAL(256, xwpp::name_to_row_2("A1:IW257"));
  BOOST_CHECK_EQUAL(16383, xwpp::name_to_row_2("A1:XFD16384"));
  BOOST_CHECK_EQUAL(16384, xwpp::name_to_row_2("A1:XFE16385"));
  BOOST_CHECK_EQUAL(1048576, xwpp::name_to_row_2("A1:XFE1048577"));
  BOOST_CHECK_EQUAL(1048576, xwpp::name_to_row_2("A1:$XFE$1048577"));
}

BOOST_AUTO_TEST_SUITE_END()
