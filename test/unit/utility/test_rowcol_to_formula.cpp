/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(rowcol_to_formula_abs)
{
  BOOST_CHECK_EQUAL("Sheet1!$A$1:$B$1", xwpp::rowcol_to_formula_abs("Sheet1", 0, 0, 0, 1));
  BOOST_CHECK_EQUAL("Sheet1!$C$1:$J$1", xwpp::rowcol_to_formula_abs("Sheet1", 0, 2, 0, 9));
  BOOST_CHECK_EQUAL("Sheet1!$A$2:$A$3", xwpp::rowcol_to_formula_abs("Sheet1", 1, 0, 2, 0));
  BOOST_CHECK_EQUAL("Sheet1!$A$10:$Y$2", xwpp::rowcol_to_formula_abs("Sheet1", 9, 0, 1, 24));
  BOOST_CHECK_EQUAL("Sheet1!$Z$8:$AA$10", xwpp::rowcol_to_formula_abs("Sheet1", 7, 25, 9, 26));
  BOOST_CHECK_EQUAL("Sheet1!$IU$2:$IV$2", xwpp::rowcol_to_formula_abs("Sheet1", 1, 254, 1, 255));
  BOOST_CHECK_EQUAL("Sheet1!$IW$2:$XFD$1", xwpp::rowcol_to_formula_abs("Sheet1", 1, 256, 0, 16383));
  BOOST_CHECK_EQUAL("Sheet1!$A$1:$XFE$1048577", xwpp::rowcol_to_formula_abs("Sheet1", 0, 0, 1048576, 16384));
  BOOST_CHECK_EQUAL("Sheet1!$XFD$1048576:$XFE$1048577",
                    xwpp::rowcol_to_formula_abs("Sheet1", 1048575, 16383, 1048576, 16384));

  BOOST_CHECK_EQUAL("'New data'!$C$2:$C$9", xwpp::rowcol_to_formula_abs("New data", 1, 2, 8, 2));
  BOOST_CHECK_EQUAL("'New data'!$C$2:$C$9", xwpp::rowcol_to_formula_abs("'New data'", 1, 2, 8, 2));

  BOOST_CHECK_EQUAL("Sheet1!$A$1", xwpp::rowcol_to_formula_abs("Sheet1", 0, 0, 0, 0));
  BOOST_CHECK_EQUAL("Sheet1!$XFE$1048577", xwpp::rowcol_to_formula_abs("Sheet1", 1048576, 16384, 1048576, 16384));
}

BOOST_AUTO_TEST_SUITE_END()
