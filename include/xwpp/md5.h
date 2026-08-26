/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 *
 * See https://www.rfc-editor.org/info/rfc1321/
 */

#ifndef XWPP_MD5_H
#define XWPP_MD5_H

#include <array>
#include <cstdint>
#include <string>
#include <vector>

namespace xwpp
{

class md5_t
{
public:
  // Build MD5 digest
  static std::array<uint8_t, 16> digest(const std::vector<uint8_t>& input);
  // Build MD5 digest and return it in a "hex string"
  static std::string digest_to_string(const std::vector<uint8_t>& input);

private:
  // Only usable through static functions hereafter
  md5_t() = default;

  [[nodiscard]] static uint32_t leftrotate(uint32_t x, uint32_t n);

  void step(const std::array<uint32_t, 16>& chunck);
  void update(const std::vector<uint8_t>& input);
  std::array<uint8_t, 16> finalize();

  static const std::array<uint32_t, 64> k_sine;
  static const std::array<uint32_t, 64> s_shift;

  uint32_t a0 = 0x67452301U;
  uint32_t b0 = 0xefcdab89U;
  uint32_t c0 = 0x98badcfeU;
  uint32_t d0 = 0x10325476U;

  size_t size = 0;
  std::array<uint8_t, 64> buffer{};
};

}

#endif
