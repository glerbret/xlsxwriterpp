/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"
#include "xwpp/workbook.h"
#include "xwpp/worksheet.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(worksheet)

namespace
{

uint32_t width_to_pixels(double width)
{
  const double max_digit_width = 7.0;
  const double padding         = 5.0;
  double pixels                = 0.;

  if(width < 1.0)
  {
    pixels = static_cast<uint32_t>((width * (max_digit_width + padding)) + 0.5);
  }
  else
  {
    pixels = static_cast<uint32_t>((width * max_digit_width + 0.5)) + 5;
  }

  return static_cast<uint32_t>(pixels);
}

uint32_t height_to_pixels(double height)
{
  return static_cast<uint32_t>(height / 0.75);
}

}

BOOST_AUTO_TEST_CASE(pixel_to_width, *boost::unit_test::tolerance(0.00001))
{
  for(int pixels = 0; pixels <= 1790; pixels++)
  {
    BOOST_TEST(pixels == width_to_pixels(xwpp::pixels_to_width(pixels)));
  }
}

BOOST_AUTO_TEST_CASE(pixel_to_height, *boost::unit_test::tolerance(0.00001))
{
  for(int pixels = 0; pixels <= 545; pixels++)
  {
    BOOST_TEST(pixels == height_to_pixels(xwpp::pixels_to_height(pixels)));
  }
}

BOOST_AUTO_TEST_SUITE_END()
