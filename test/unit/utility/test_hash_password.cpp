/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(hash_password)
{
  BOOST_CHECK_EQUAL(0x83AF, xwpp::hash_password("password"));
  BOOST_CHECK_EQUAL(0xD14E, xwpp::hash_password("This is a longer phrase"));
  BOOST_CHECK_EQUAL(0xCE2A, xwpp::hash_password("0"));
  BOOST_CHECK_EQUAL(0xCEED, xwpp::hash_password("01"));
  BOOST_CHECK_EQUAL(0xCF7C, xwpp::hash_password("012"));
  BOOST_CHECK_EQUAL(0xCC4B, xwpp::hash_password("0123"));
  BOOST_CHECK_EQUAL(0xCACA, xwpp::hash_password("01234"));
  BOOST_CHECK_EQUAL(0xC789, xwpp::hash_password("012345"));
  BOOST_CHECK_EQUAL(0xDC88, xwpp::hash_password("0123456"));
  BOOST_CHECK_EQUAL(0xEB87, xwpp::hash_password("01234567"));
  BOOST_CHECK_EQUAL(0x9B86, xwpp::hash_password("012345678"));
  BOOST_CHECK_EQUAL(0xFF84, xwpp::hash_password("0123456789"));
  BOOST_CHECK_EQUAL(0xFF86, xwpp::hash_password("01234567890"));
  BOOST_CHECK_EQUAL(0xEF87, xwpp::hash_password("012345678901"));
  BOOST_CHECK_EQUAL(0xAF8A, xwpp::hash_password("0123456789012"));
  BOOST_CHECK_EQUAL(0xEF90, xwpp::hash_password("01234567890123"));
  BOOST_CHECK_EQUAL(0xEFA5, xwpp::hash_password("012345678901234"));
  BOOST_CHECK_EQUAL(0xEFD0, xwpp::hash_password("0123456789012345"));
  BOOST_CHECK_EQUAL(0xEF09, xwpp::hash_password("01234567890123456"));
  BOOST_CHECK_EQUAL(0xEEB2, xwpp::hash_password("012345678901234567"));
  BOOST_CHECK_EQUAL(0xED33, xwpp::hash_password("0123456789012345678"));
  BOOST_CHECK_EQUAL(0xEA14, xwpp::hash_password("01234567890123456789"));
  BOOST_CHECK_EQUAL(0xE615, xwpp::hash_password("012345678901234567890"));
  BOOST_CHECK_EQUAL(0xFE96, xwpp::hash_password("0123456789012345678901"));
  BOOST_CHECK_EQUAL(0xCC97, xwpp::hash_password("01234567890123456789012"));
  BOOST_CHECK_EQUAL(0xAA98, xwpp::hash_password("012345678901234567890123"));
  BOOST_CHECK_EQUAL(0xFA98, xwpp::hash_password("0123456789012345678901234"));
  BOOST_CHECK_EQUAL(0xD298, xwpp::hash_password("01234567890123456789012345"));
  BOOST_CHECK_EQUAL(0xD2D3, xwpp::hash_password("0123456789012345678901234567890"));
}

BOOST_AUTO_TEST_SUITE_END()
