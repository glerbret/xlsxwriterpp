/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/md5.h"

#include <array>
#include <cstdint>
#include <format>
#include <string>
#include <vector>

namespace xwpp
{

const std::array<uint32_t, 64> md5_t::k_sine{
  0xd76aa478, 0xe8c7b756, 0x242070db, 0xc1bdceee, 0xf57c0faf, 0x4787c62a, 0xa8304613, 0xfd469501,
  0x698098d8, 0x8b44f7af, 0xffff5bb1, 0x895cd7be, 0x6b901122, 0xfd987193, 0xa679438e, 0x49b40821,
  0xf61e2562, 0xc040b340, 0x265e5a51, 0xe9b6c7aa, 0xd62f105d, 0x02441453, 0xd8a1e681, 0xe7d3fbc8,
  0x21e1cde6, 0xc33707d6, 0xf4d50d87, 0x455a14ed, 0xa9e3e905, 0xfcefa3f8, 0x676f02d9, 0x8d2a4c8a,
  0xfffa3942, 0x8771f681, 0x6d9d6122, 0xfde5380c, 0xa4beea44, 0x4bdecfa9, 0xf6bb4b60, 0xbebfbc70,
  0x289b7ec6, 0xeaa127fa, 0xd4ef3085, 0x04881d05, 0xd9d4d039, 0xe6db99e5, 0x1fa27cf8, 0xc4ac5665,
  0xf4292244, 0x432aff97, 0xab9423a7, 0xfc93a039, 0x655b59c3, 0x8f0ccc92, 0xffeff47d, 0x85845dd1,
  0x6fa87e4f, 0xfe2ce6e0, 0xa3014314, 0x4e0811a1, 0xf7537e82, 0xbd3af235, 0x2ad7d2bb, 0xeb86d391};

const std::array<uint32_t, 64> md5_t::s_shift{7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22,
                                              5, 9,  14, 20, 5, 9,  14, 20, 5, 9,  14, 20, 5, 9,  14, 20,
                                              4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23,
                                              6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21};

uint32_t md5_t::leftrotate(uint32_t x, uint32_t n)
{
  return (x << n) | (x >> (32 - n));
}

// Process one 512-bit chunks.
void md5_t::step(const std::array<uint32_t, 16>& chunck)
{
  uint32_t a = a0;
  uint32_t b = b0;
  uint32_t c = c0;
  uint32_t d = d0;

  for(uint32_t i = 0; i < 64; ++i)
  {
    uint32_t f = 0;
    uint32_t g = 0;

    switch(i / 16)
    {
      case 0:
        f = (b & c) | (~b & d);
        g = i;
        break;

      case 1:
        f = (b & d) | (c & ~d);
        g = ((i * 5) + 1) % 16;
        break;

      case 2:
        f = b ^ c ^ d;
        g = ((i * 3) + 5) % 16;
        break;

      default:
        f = c ^ (b | ~d);
        g = (i * 7) % 16;
        break;
    }

    // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
    f = f + a + k_sine[i] + chunck[g];
    a = d;
    d = c;
    c = b;
    // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
    b = b + leftrotate(f, s_shift[i]);
  }

  a0 += a;
  b0 += b;
  c0 += c;
  d0 += d;
}

void md5_t::update(const std::vector<uint8_t>& input)
{
  size_t offset = size % 64;
  size += input.size();

  for(const auto byte: input)
  {
    // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
    buffer[offset++] = byte;

    // 512-bits chunk available, call step
    if(offset % 64 == 0)
    {
      std::array<uint32_t, 16> chunk{};

      for(size_t i = 0; i < chunk.size(); ++i)
      {
        // NOLINTBEGIN(cppcoreguidelines-pro-bounds-constant-array-index)
        chunk[i] = static_cast<uint32_t>(buffer[(i * 4) + 3]) << 24U |
                   static_cast<uint32_t>(buffer[(i * 4) + 2]) << 16U |
                   static_cast<uint32_t>(buffer[(i * 4) + 1]) << 8U | static_cast<uint32_t>(buffer[(i * 4)]);
        // NOLINTEND(cppcoreguidelines-pro-bounds-constant-array-index)
      }

      step(chunk);
      offset = 0;
    }
  }
}

std::array<uint8_t, 16> md5_t::finalize()
{
  const size_t offset = size % 64;

  // Pad the input (always present)
  const size_t padding_len = offset < 56 ? 56 - offset : 120 - offset;
  std::vector<uint8_t> padding(padding_len, 0x00);
  padding[0] = 0x80;
  update(padding);

  // Process last block with two last words containing the input size
  size -= padding_len;
  std::array<uint32_t, 16> chunk{};
  for(size_t i = 0; i < 14; ++i)
  {
    // NOLINTBEGIN(cppcoreguidelines-pro-bounds-constant-array-index)
    chunk[i] = static_cast<uint32_t>(buffer[(i * 4) + 3]) << 24U | static_cast<uint32_t>(buffer[(i * 4) + 2]) << 16U |
               static_cast<uint32_t>(buffer[(i * 4) + 1]) << 8U | static_cast<uint32_t>(buffer[(i * 4)]);
    // NOLINTEND(cppcoreguidelines-pro-bounds-constant-array-index)
  }
  chunk[14] = static_cast<uint32_t>(size * 8);
  chunk[15] = static_cast<uint32_t>((size * 8) >> 32U);
  step(chunk);

  return std::array<uint8_t, 16>{
    static_cast<uint8_t>(a0 & 0x000000ffU),          static_cast<uint8_t>((a0 & 0x0000ff00U) >> 8U),
    static_cast<uint8_t>((a0 & 0x00ff0000U) >> 16U), static_cast<uint8_t>((a0 & 0xff000000U) >> 24U),
    static_cast<uint8_t>(b0 & 0x000000ffU),          static_cast<uint8_t>((b0 & 0x0000ff00U) >> 8U),
    static_cast<uint8_t>((b0 & 0x00ff0000U) >> 16U), static_cast<uint8_t>((b0 & 0xff000000U) >> 24U),
    static_cast<uint8_t>(c0 & 0x000000ffU),          static_cast<uint8_t>((c0 & 0x0000ff00U) >> 8U),
    static_cast<uint8_t>((c0 & 0x00ff0000U) >> 16U), static_cast<uint8_t>((c0 & 0xff000000U) >> 24U),
    static_cast<uint8_t>(d0 & 0x000000ffU),          static_cast<uint8_t>((d0 & 0x0000ff00U) >> 8U),
    static_cast<uint8_t>((d0 & 0x00ff0000U) >> 16U), static_cast<uint8_t>((d0 & 0xff000000U) >> 24U),
  };
}

std::array<uint8_t, 16> md5_t::digest(const std::vector<uint8_t>& input)
{
  md5_t md5;
  md5.update(input);
  return md5.finalize();
}

std::string md5_t::digest_to_string(const std::vector<uint8_t>& input)
{
  auto hash = md5_t::digest(input);

  std::string hash_str;
  for(const auto b: hash)
  {
    hash_str += std::format("{:02X}", b);
  }

  return hash_str;
}

}
