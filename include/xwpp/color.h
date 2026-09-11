/*
 * chartsheet - A Xlsxwriter++ library for creating Excel XLSX chartsheet files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @page color_page The object used to manage colors.
 *
 * @see @ref color.h for full details of the functionality.
 *
 * @file
 *
 * @brief Object used to manage colors.
 *
 * Object used to manage colors.
 */

#ifndef XWPP_COLOR_H
#define XWPP_COLOR_H

#include <cstdint>
#include <format>

/// @cond
namespace xwpp
{
/// @endcond

/**
 * @brief The type for colors in `Xlsxwriter++`.
 *
 * The type for colors in `Xlsxwriter++`.
 *
 * @see @ref working_with_colors.
 *
 * Few predefined colors are provided.
 */
class color_t
{
public:
  color_t() = default;

  /**
   * @brief Create color from HTML RGB integer value.
   *
   * @param color The HTML RGB color integer value.
   *
   * The `%color_t()` constructor creates a color from HTML RGB integer value.
   *
   * @code
   *  xwpp::color_t{0xFF9900};  // Orange.
   * @endcode
   *
   * @pre color <= 0XFFFFFF
   */
  explicit color_t(uint32_t color);

  /**
   * @brief Create color from RGB component.
   *
   * @param red   The red part of RGB color.
   * @param green The green part of RGB color.
   * @param blue  The blue part of RGB color.
   *
   * The `%rgb()` function creates a color from RGB component.
   *
   * @code
   *  xwpp::color_t::rgb(0xFF, 0x99, 0x00);  // Orange.
   * @endcode
   */
  static color_t rgb(uint8_t red, uint8_t green, uint8_t blue);

  /**
   * @brief Create color from CMYK percents.
   *
   * @param cyan    The cyan part of CMYK color.
   * @param magenta The magenta part of CMYK color.
   * @param yellow  The yellow part of CMYK color.
   * @param black   The black part of RGB color.
   *
   * The `%cmyk()` function creates a color from CYMK percentage.
   *
   * @code
   *  xwpp::color_t::cmyk(0, 40, 100, 0);  // Orange.
   * @endcode
   *
   * @pre cyan <= 100
   * @pre magenta <= 100
   * @pre yellow <= 100
   * @pre black <= 100
   */
  static color_t cmyk(uint8_t cyan, uint8_t magenta, uint8_t yellow, uint8_t black);

  explicit operator bool() const
  {
    return color_ != UNSET;
  }

  friend bool operator==(const color_t& lhs, const color_t& rhs) = default;
  friend bool operator!=(const color_t& lhs, const color_t& rhs) = default;

  /**
   * @brief Predefined color black.
   */
  static color_t black();
  /**
   * @brief Predefined color blue.
   */
  static color_t blue();
  /**
   * @brief Predefined color brown.
   */
  static color_t brown();
  /**
   * @brief Predefined color cyan.
   */
  static color_t cyan();
  /**
   * @brief Predefined color gray.
   */
  static color_t gray();
  /**
   * @brief Predefined color green.
   */
  static color_t green();
  /**
   * @brief Predefined color lime.
   */
  static color_t lime();
  /**
   * @brief Predefined color magenta.
   */
  static color_t magenta();
  /**
   * @brief Predefined color navy.
   */
  static color_t navy();
  /**
   * @brief Predefined color orange.
   */
  static color_t orange();
  /**
   * @brief Predefined color pink.
   */
  static color_t pink();
  /**
   * @brief Predefined color purple.
   */
  static color_t purple();
  /**
   * @brief Predefined color red.
   */
  static color_t red();
  /**
   * @brief Predefined color silver.
   */
  static color_t silver();
  /**
   * @brief Predefined color white.
   */
  static color_t white();
  /**
   * @brief Predefined color yellow.
   */
  static color_t yellow();

private:
  friend struct std::formatter<xwpp::color_t>;

  static const uint32_t UNSET = 0xFFFFFFFF;
  static const uint32_t MASK  = 0x00FFFFFF;

  uint32_t color_ = UNSET;
};

/// @cond
}
/// @endcond

/// @cond
template<>
class std::formatter<xwpp::color_t>
{
public:
  constexpr auto parse(std::format_parse_context& parse_context)
  {
    // NOLINTNEXTLINE(llvm-qualified-auto,readability-qualified-auto)
    auto pos = parse_context.begin();
    while(pos != parse_context.end() && *pos != '}')
    {
      if(*pos == 'p' || *pos == 'P')
      {
        padding_ = true;
      }

      if(*pos == 'P' || *pos == 'X')
      {
        uppercase_ = true;
      }
      ++pos;
    }
    return pos;
  }

  auto format(const xwpp::color_t& data, std::format_context& format_context) const
  {
    std::string out;
    if(uppercase_)
    {
      out = std::format("{}{:06X}", padding_ ? "FF" : "", data.color_ & xwpp::color_t::MASK);
    }
    else
    {
      out = std::format("{}{:06x}", padding_ ? "FF" : "", data.color_ & xwpp::color_t::MASK);
    }
    return formatter_.format(out, format_context);
  }

private:
  std::formatter<std::string> formatter_;
  bool padding_   = false;
  bool uppercase_ = false;
};
/// @endcond

#endif
