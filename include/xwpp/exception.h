/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @file
 *
 * @brief Exceptions of Xlsxwriter++.
 *
 * Exceptions of Xlsxwriter++.
 */

#ifndef XWPP_EXCEPTION_H
#define XWPP_EXCEPTION_H

#include <stdexcept>
#include <string>

/// @cond
namespace xwpp
{
/// @endcond

/**
 * @brief Base exception of `Xlsxwriter++`.
 *
 * Base exception of `Xlsxwriter++`.
 */
class xwpp_exception_t : public std::runtime_error
{
public:
  /**
   * @param what_arg Description of the exception.
   */
  explicit xwpp_exception_t(const std::string& what_arg)
    : std::runtime_error(what_arg)
  {
  }

  /// @overload
  explicit xwpp_exception_t(const char* what_arg)
    : std::runtime_error(what_arg)
  {
  }
};

/**
 * @brief Out of range exception.
 *
 * Out of range exception.
 */
class xwpp_out_of_range_t : public xwpp_exception_t
{
public:
  /**
   * @param what_arg Description of the exception.
   */
  explicit xwpp_out_of_range_t(const std::string& what_arg)
    : xwpp_exception_t(what_arg)
  {
  }

  /// @overload
  explicit xwpp_out_of_range_t(const char* what_arg)
    : xwpp_exception_t(what_arg)
  {
  }
};

/// @cond
}
/// @endcond

#endif
