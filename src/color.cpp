/*
 * utility - Utility functions for Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/color.h"

#include "xlsxwriterpp.h"

#include <cassert>

namespace xwpp
{

color_t::color_t(uint32_t color)
  : color_{color}
{
  assert(color <= 0xFFFFFF);
}

color_t color_t::rgb(uint8_t red, uint8_t green, uint8_t blue)
{
  return color_t{static_cast<uint32_t>(red * 0x10000 + green * 0x100 + blue)};
}

color_t color_t::cmyk(uint8_t cyan, uint8_t magenta, uint8_t yellow, uint8_t black)
{
  assert(cyan <= 100);
  assert(magenta <= 100);
  assert(yellow <= 100);
  assert(black <= 100);

  return color_t{
    static_cast<uint32_t>(static_cast<uint8_t>(255 * (100. - cyan) / 100. * (100. - black) / 100.) * 0x10000 +
                          static_cast<uint8_t>(255 * (100. - magenta) / 100. * (100. - black) / 100.) * 0x100 +
                          static_cast<uint8_t>(255 * (100. - yellow) / 100. * (100. - black) / 100.))};
}

color_t color_t::black()
{
  return color_t{0x000000};
}

color_t color_t::blue()
{
  return color_t{0x0000FF};
}

color_t color_t::brown()
{
  return color_t{0x800000};
}

color_t color_t::cyan()
{
  return color_t{0x00FFFF};
}

color_t color_t::gray()
{
  return color_t{0x808080};
}

color_t color_t::green()
{
  return color_t{0x008000};
}

color_t color_t::lime()
{
  return color_t{0x00FF00};
}

color_t color_t::magenta()
{
  return color_t{0xFF00FF};
}

color_t color_t::navy()
{
  return color_t{0x000080};
}

color_t color_t::orange()
{
  return color_t{0xFF6600};
}

color_t color_t::pink()
{
  return color_t{0xFF00FF};
}

color_t color_t::purple()
{
  return color_t{0x800080};
}

color_t color_t::red()
{
  return color_t{0xFF0000};
}

color_t color_t::silver()
{
  return color_t{0xC0C0C0};
}

color_t color_t::white()
{
  return color_t{0xFFFFFF};
}

color_t color_t::yellow()
{
  return color_t{0xFFFF00};
}

}