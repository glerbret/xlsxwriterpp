/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

#include <chrono>

using namespace std::literals::chrono_literals;

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(datetime_to_excel_datetime_chrono, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(
    0 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{1899y / std::chrono::December / 31d} + 0h + 0min + 0s));
  BOOST_TEST(30188.010650613425 ==
             xwpp::datetime_to_excel_datetime(std::chrono::sys_days{1982y / std::chrono::August / 25d} + 0h + 15min +
                                              20s + 213ms));
  BOOST_TEST(60376.011670023145 ==
             xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2065y / std::chrono::April / 19d} + 0h + 16min +
                                              48s + 290ms));
  BOOST_TEST(90565.038488958337 ==
             xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2147y / std::chrono::December / 15d} + 0h + 55min +
                                              25s + 446ms));
  BOOST_TEST(120753.04359827546 ==
             xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2230y / std::chrono::August / 10d} + 1h + 2min +
                                              46s + 891ms));
}

BOOST_AUTO_TEST_CASE(datetime_date_only_chrono, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{1899y / std::chrono::December / 31d}));
  BOOST_TEST(1461 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{1903y / std::chrono::December / 31d}));
  BOOST_TEST(36708 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2000y / std::chrono::July / 1d}));
  BOOST_TEST(60376 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2065y / std::chrono::April / 19d}));
  BOOST_TEST(120753 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2230y / std::chrono::August / 10d}));
}

BOOST_AUTO_TEST_CASE(datetime_date_only_1904_chrono, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 ==
             xwpp::datetime_to_excel_date_with_epoch(std::chrono::sys_days{1904y / std::chrono::January / 1d}, true));
  BOOST_TEST(1159 ==
             xwpp::datetime_to_excel_date_with_epoch(std::chrono::sys_days{1907y / std::chrono::March / 5d}, true));
  BOOST_TEST(34879 ==
             xwpp::datetime_to_excel_date_with_epoch(std::chrono::sys_days{1999y / std::chrono::June / 30d}, true));
  BOOST_TEST(35702 == xwpp::datetime_to_excel_date_with_epoch(
                        std::chrono::sys_days{2001y / std::chrono::September / 30d}, true));
  BOOST_TEST(58914 ==
             xwpp::datetime_to_excel_date_with_epoch(std::chrono::sys_days{2065y / std::chrono::April / 19d}, true));
}

BOOST_AUTO_TEST_CASE(datetime_time_only_chrono, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{0h + 0min + 0s}));
  BOOST_TEST(0.11472348379629631 ==
             xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{2h + 45min + 12s + 109ms}));
  BOOST_TEST(0.39631690972222228 ==
             xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{9h + 30min + 41s + 781ms}));
  BOOST_TEST(0.6874536921296297 ==
             xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{16h + 29min + 55s + 999ms}));
  BOOST_TEST(0.90703436342592603 ==
             xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{21h + 46min + 7s + 769ms}));
}

BOOST_AUTO_TEST_CASE(datetime_to_excel_datetime, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1899, 12, 31, 0, 0, 0.000}));
  BOOST_TEST(30188.010650613425 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1982, 8, 25, 0, 15, 20.213}));
  BOOST_TEST(60376.011670023145 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2065, 4, 19, 0, 16, 48.290}));
  BOOST_TEST(90565.038488958337 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2147, 12, 15, 0, 55, 25.446}));
  BOOST_TEST(120753.04359827546 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2230, 8, 10, 1, 2, 46.891}));
  BOOST_TEST(150942.04462496529 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2313, 4, 6, 1, 4, 15.597}));
  BOOST_TEST(181130.04838991899 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2395, 11, 30, 1, 9, 40.889}));
  BOOST_TEST(211318.04968240741 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2478, 7, 25, 1, 11, 32.560}));
  BOOST_TEST(241507.06272186342 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2561, 3, 21, 1, 30, 19.169}));
  BOOST_TEST(271695.07529606484 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2643, 11, 15, 1, 48, 25.580}));
  BOOST_TEST(301884.08578609955 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2726, 7, 12, 2, 3, 31.919}));
  BOOST_TEST(332072.09111094906 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2809, 3, 6, 2, 11, 11.986}));
  BOOST_TEST(362261.10042934027 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2891, 10, 31, 2, 24, 37.095}));
  BOOST_TEST(392449.10772245371 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2974, 6, 26, 2, 35, 7.220}));
  BOOST_TEST(422637.1147234838 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3057, 2, 19, 2, 45, 12.109}));
  BOOST_TEST(452826.12962951389 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3139, 10, 17, 3, 6, 39.990}));
  BOOST_TEST(483014.13065105322 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3222, 6, 11, 3, 8, 8.251}));
  BOOST_TEST(513203.13834 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3305, 2, 5, 3, 19, 12.576}));
  BOOST_TEST(543391.14563164348 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3387, 10, 1, 3, 29, 42.574}));
  BOOST_TEST(573579.15105107636 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3470, 5, 27, 3, 37, 30.813}));
  BOOST_TEST(603768.17683137732 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3553, 1, 21, 4, 14, 38.231}));
  BOOST_TEST(633956.17810832174 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3635, 9, 16, 4, 16, 28.559}));
  BOOST_TEST(664145.17914608796 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3718, 5, 13, 4, 17, 58.222}));
  BOOST_TEST(694333.18173372687 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3801, 1, 6, 4, 21, 41.794}));
  BOOST_TEST(724522.20596981479 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3883, 9, 2, 4, 56, 35.792}));
  BOOST_TEST(754710.2258667245 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{3966, 4, 28, 5, 25, 14.885}));
  BOOST_TEST(784898.22645513888 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4048, 12, 21, 5, 26, 5.724}));
  BOOST_TEST(815087.24078782403 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4131, 8, 18, 5, 46, 44.068}));
  BOOST_TEST(845275.24167987274 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4214, 4, 13, 5, 48, 1.141}));
  BOOST_TEST(875464.24574438657 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4296, 12, 7, 5, 53, 52.315}));
  BOOST_TEST(905652.26028449077 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4379, 8, 3, 6, 14, 48.580}));
  BOOST_TEST(935840.28212659725 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4462, 3, 28, 6, 46, 15.738}));
  BOOST_TEST(966029.31343063654 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4544, 11, 22, 7, 31, 20.407}));
  BOOST_TEST(996217.33233511576 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4627, 7, 19, 7, 58, 33.754}));
  BOOST_TEST(1026406.3386936343 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4710, 3, 15, 8, 7, 43.130}));
  BOOST_TEST(1056594.3536005903 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4792, 11, 7, 8, 29, 11.091}));
  BOOST_TEST(1086783.3807329629 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4875, 7, 4, 9, 8, 15.328}));
  BOOST_TEST(1116971.3963169097 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4958, 2, 27, 9, 30, 41.781}));
  BOOST_TEST(1147159.3986627546 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5040, 10, 23, 9, 34, 4.462}));
  BOOST_TEST(1177348.4009715857 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5123, 6, 20, 9, 37, 23.945}));
  BOOST_TEST(1207536.4013501736 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5206, 2, 12, 9, 37, 56.655}));
  BOOST_TEST(1237725.406391551 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5288, 10, 8, 9, 45, 12.230}));
  BOOST_TEST(1267913.412671088 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5371, 6, 4, 9, 54, 14.782}));
  BOOST_TEST(1298101.4127558796 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5454, 1, 28, 9, 54, 22.108}));
  BOOST_TEST(1328290.4177795255 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5536, 9, 24, 10, 1, 36.151}));
  BOOST_TEST(1358478.5068125231 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5619, 5, 20, 12, 9, 48.602}));
  BOOST_TEST(1388667.5237100578 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5702, 1, 14, 12, 34, 8.549}));
  BOOST_TEST(1418855.5389640625 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5784, 9, 8, 12, 56, 6.495}));
  BOOST_TEST(1449044.5409515856 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5867, 5, 6, 12, 58, 58.217}));
  BOOST_TEST(1479232.5416002662 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{5949, 12, 30, 12, 59, 54.263}));
  BOOST_TEST(1509420.5657561459 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6032, 8, 24, 13, 34, 41.331}));
  BOOST_TEST(1539609.5822754744 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6115, 4, 21, 13, 58, 28.601}));
  BOOST_TEST(1569797.5849178126 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6197, 12, 14, 14, 2, 16.899}));
  BOOST_TEST(1599986.6085352316 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6280, 8, 10, 14, 36, 17.444}));
  BOOST_TEST(1630174.60969272 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6363, 4, 6, 14, 37, 57.451}));
  BOOST_TEST(1660363.6234115392 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6445, 11, 30, 14, 57, 42.757}));
  BOOST_TEST(1690551.6325035533 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6528, 7, 26, 15, 10, 48.307}));
  BOOST_TEST(1720739.635183912 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6611, 3, 22, 15, 14, 39.890}));
  BOOST_TEST(1750928.6387498612 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6693, 11, 15, 15, 19, 47.988}));
  BOOST_TEST(1781116.6697262037 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6776, 7, 11, 16, 4, 24.344}));
  BOOST_TEST(1811305.6822216667 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6859, 3, 7, 16, 22, 23.952}));
  BOOST_TEST(1841493.6874536921 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{6941, 10, 31, 16, 29, 55.999}));
  BOOST_TEST(1871681.7071789235 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7024, 6, 26, 16, 58, 20.259}));
  BOOST_TEST(1901870.7111390624 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7107, 2, 21, 17, 4, 2.415}));
  BOOST_TEST(1932058.7211762732 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7189, 10, 16, 17, 18, 29.630}));
  BOOST_TEST(1962247.7412190163 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7272, 6, 11, 17, 47, 21.323}));
  BOOST_TEST(1992435.7454845603 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7355, 2, 5, 17, 53, 29.866}));
  BOOST_TEST(2022624.7456143056 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7437, 10, 2, 17, 53, 41.076}));
  BOOST_TEST(2052812.7465977315 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7520, 5, 28, 17, 55, 6.044}));
  BOOST_TEST(2083000.7602910995 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7603, 1, 21, 18, 14, 49.151}));
  BOOST_TEST(2113189.7623349307 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7685, 9, 16, 18, 17, 45.738}));
  BOOST_TEST(2143377.7708298611 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7768, 5, 12, 18, 29, 59.700}));
  BOOST_TEST(2173566.773162419 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7851, 1, 7, 18, 33, 21.233}));
  BOOST_TEST(2203754.8016744559 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{7933, 9, 2, 19, 14, 24.673}));
  BOOST_TEST(2233942.8036205554 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8016, 4, 27, 19, 17, 12.816}));
  BOOST_TEST(2264131.8080603937 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8098, 12, 22, 19, 23, 36.418}));
  BOOST_TEST(2294319.8239109721 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8181, 8, 17, 19, 46, 25.908}));
  BOOST_TEST(2324508.8387420601 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8264, 4, 13, 20, 7, 47.314}));
  BOOST_TEST(2354696.855296331 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8346, 12, 8, 20, 31, 37.603}));
  BOOST_TEST(2384885.8610853008 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8429, 8, 3, 20, 39, 57.770}));
  BOOST_TEST(2415073.8682530904 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8512, 3, 29, 20, 50, 17.067}));
  BOOST_TEST(2445261.8770581828 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8594, 11, 22, 21, 2, 57.827}));
  BOOST_TEST(2475450.8910360998 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8677, 7, 19, 21, 23, 5.519}));
  BOOST_TEST(2505638.8991848612 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8760, 3, 14, 21, 34, 49.572}));
  BOOST_TEST(2535827.9021521294 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8842, 11, 8, 21, 39, 5.944}));
  BOOST_TEST(2566015.9022965971 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{8925, 7, 4, 21, 39, 18.426}));
  BOOST_TEST(2596203.9070343636 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9008, 2, 28, 21, 46, 7.769}));
  BOOST_TEST(2626392.9152275696 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9090, 10, 24, 21, 57, 55.662}));
  BOOST_TEST(2656580.9299968979 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9173, 6, 19, 22, 19, 11.732}));
  BOOST_TEST(2686769.9332335186 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9256, 2, 13, 22, 23, 51.376}));
  BOOST_TEST(2716957.9360968866 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9338, 10, 9, 22, 27, 58.771}));
  BOOST_TEST(2747146.9468795368 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9421, 6, 5, 22, 43, 30.392}));
  BOOST_TEST(2777334.9502990046 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9504, 1, 30, 22, 48, 25.834}));
  BOOST_TEST(2807522.9540709145 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9586, 9, 24, 22, 53, 51.727}));
  BOOST_TEST(2837711.9673210187 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9669, 5, 20, 23, 12, 56.536}));
  BOOST_TEST(2867899.9693762613 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9752, 1, 14, 23, 15, 54.109}));
  BOOST_TEST(2898088.9702850925 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9834, 9, 10, 23, 17, 12.632}));
  BOOST_TEST(2958465.999988426 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 12, 31, 23, 59, 59.000}));
}

BOOST_AUTO_TEST_CASE(datetime_to_excel_datetime_date_only, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1899, 12, 31}));
  BOOST_TEST(1 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1900, 1, 1}));
  BOOST_TEST(58 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1900, 2, 27}));
  BOOST_TEST(59 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1900, 2, 28}));
  BOOST_TEST(61 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1900, 3, 1}));
  BOOST_TEST(62 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1900, 3, 2}));
  BOOST_TEST(71 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1900, 3, 11}));
  BOOST_TEST(99 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1900, 4, 8}));
  BOOST_TEST(256 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1900, 9, 12}));
  BOOST_TEST(489 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1901, 5, 3}));
  BOOST_TEST(652 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1901, 10, 13}));
  BOOST_TEST(777 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1902, 2, 15}));
  BOOST_TEST(888 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1902, 6, 6}));
  BOOST_TEST(999 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1902, 9, 25}));
  BOOST_TEST(1001 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1902, 9, 27}));
  BOOST_TEST(1212 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1903, 4, 26}));
  BOOST_TEST(1313 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1903, 8, 5}));
  BOOST_TEST(1461 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1903, 12, 31}));
  BOOST_TEST(1462 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1904, 1, 1}));
  BOOST_TEST(1520 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1904, 2, 28}));
  BOOST_TEST(1521 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1904, 2, 29}));
  BOOST_TEST(1522 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1904, 3, 1}));
  BOOST_TEST(2615 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1907, 2, 27}));
  BOOST_TEST(2616 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1907, 2, 28}));
  BOOST_TEST(2617 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1907, 3, 1}));
  BOOST_TEST(2618 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1907, 3, 2}));
  BOOST_TEST(2619 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1907, 3, 3}));
  BOOST_TEST(2620 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1907, 3, 4}));
  BOOST_TEST(2621 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1907, 3, 5}));
  BOOST_TEST(2622 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1907, 3, 6}));
  BOOST_TEST(36161 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 1, 1}));
  BOOST_TEST(36191 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 1, 31}));
  BOOST_TEST(36192 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 2, 1}));
  BOOST_TEST(36219 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 2, 28}));
  BOOST_TEST(36220 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 3, 1}));
  BOOST_TEST(36250 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 3, 31}));
  BOOST_TEST(36251 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 4, 1}));
  BOOST_TEST(36280 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 4, 30}));
  BOOST_TEST(36281 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 5, 1}));
  BOOST_TEST(36311 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 5, 31}));
  BOOST_TEST(36312 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 6, 1}));
  BOOST_TEST(36341 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 6, 30}));
  BOOST_TEST(36342 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 7, 1}));
  BOOST_TEST(36372 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 7, 31}));
  BOOST_TEST(36373 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 8, 1}));
  BOOST_TEST(36403 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 8, 31}));
  BOOST_TEST(36404 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 9, 1}));
  BOOST_TEST(36433 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 9, 30}));
  BOOST_TEST(36434 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 10, 1}));
  BOOST_TEST(36464 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 10, 31}));
  BOOST_TEST(36465 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 11, 1}));
  BOOST_TEST(36494 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 11, 30}));
  BOOST_TEST(36495 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 12, 1}));
  BOOST_TEST(36525 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{1999, 12, 31}));
  BOOST_TEST(36526 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 1, 1}));
  BOOST_TEST(36556 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 1, 31}));
  BOOST_TEST(36557 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 2, 1}));
  BOOST_TEST(36585 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 2, 29}));
  BOOST_TEST(36586 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 3, 1}));
  BOOST_TEST(36616 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 3, 31}));
  BOOST_TEST(36617 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 4, 1}));
  BOOST_TEST(36646 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 4, 30}));
  BOOST_TEST(36647 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 5, 1}));
  BOOST_TEST(36677 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 5, 31}));
  BOOST_TEST(36678 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 6, 1}));
  BOOST_TEST(36707 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 6, 30}));
  BOOST_TEST(36708 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 7, 1}));
  BOOST_TEST(36738 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 7, 31}));
  BOOST_TEST(36739 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 8, 1}));
  BOOST_TEST(36769 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 8, 31}));
  BOOST_TEST(36770 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 9, 1}));
  BOOST_TEST(36799 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 9, 30}));
  BOOST_TEST(36800 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 10, 1}));
  BOOST_TEST(36830 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 10, 31}));
  BOOST_TEST(36831 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 11, 1}));
  BOOST_TEST(36860 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 11, 30}));
  BOOST_TEST(36861 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 12, 1}));
  BOOST_TEST(36891 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2000, 12, 31}));
  BOOST_TEST(36892 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 1, 1}));
  BOOST_TEST(36922 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 1, 31}));
  BOOST_TEST(36923 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 2, 1}));
  BOOST_TEST(36950 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 2, 28}));
  BOOST_TEST(36951 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 3, 1}));
  BOOST_TEST(36981 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 3, 31}));
  BOOST_TEST(36982 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 4, 1}));
  BOOST_TEST(37011 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 4, 30}));
  BOOST_TEST(37012 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 5, 1}));
  BOOST_TEST(37042 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 5, 31}));
  BOOST_TEST(37043 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 6, 1}));
  BOOST_TEST(37072 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 6, 30}));
  BOOST_TEST(37073 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 7, 1}));
  BOOST_TEST(37103 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 7, 31}));
  BOOST_TEST(37104 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 8, 1}));
  BOOST_TEST(37134 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 8, 31}));
  BOOST_TEST(37135 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 9, 1}));
  BOOST_TEST(37164 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 9, 30}));
  BOOST_TEST(37165 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 10, 1}));
  BOOST_TEST(37195 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 10, 31}));
  BOOST_TEST(37196 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 11, 1}));
  BOOST_TEST(37225 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 11, 30}));
  BOOST_TEST(37226 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 12, 1}));
  BOOST_TEST(37256 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2001, 12, 31}));
  BOOST_TEST(182623 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 1, 1}));
  BOOST_TEST(182653 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 1, 31}));
  BOOST_TEST(182654 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 2, 1}));
  BOOST_TEST(182682 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 2, 29}));
  BOOST_TEST(182683 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 3, 1}));
  BOOST_TEST(182713 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 3, 31}));
  BOOST_TEST(182714 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 4, 1}));
  BOOST_TEST(182743 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 4, 30}));
  BOOST_TEST(182744 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 5, 1}));
  BOOST_TEST(182774 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 5, 31}));
  BOOST_TEST(182775 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 6, 1}));
  BOOST_TEST(182804 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 6, 30}));
  BOOST_TEST(182805 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 7, 1}));
  BOOST_TEST(182835 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 7, 31}));
  BOOST_TEST(182836 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 8, 1}));
  BOOST_TEST(182866 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 8, 31}));
  BOOST_TEST(182867 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 9, 1}));
  BOOST_TEST(182896 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 9, 30}));
  BOOST_TEST(182897 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 10, 1}));
  BOOST_TEST(182927 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 10, 31}));
  BOOST_TEST(182928 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 11, 1}));
  BOOST_TEST(182957 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 11, 30}));
  BOOST_TEST(182958 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 12, 1}));
  BOOST_TEST(182988 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{2400, 12, 31}));
  BOOST_TEST(767011 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 1, 1}));
  BOOST_TEST(767041 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 1, 31}));
  BOOST_TEST(767042 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 2, 1}));
  BOOST_TEST(767070 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 2, 29}));
  BOOST_TEST(767071 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 3, 1}));
  BOOST_TEST(767101 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 3, 31}));
  BOOST_TEST(767102 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 4, 1}));
  BOOST_TEST(767131 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 4, 30}));
  BOOST_TEST(767132 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 5, 1}));
  BOOST_TEST(767162 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 5, 31}));
  BOOST_TEST(767163 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 6, 1}));
  BOOST_TEST(767192 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 6, 30}));
  BOOST_TEST(767193 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 7, 1}));
  BOOST_TEST(767223 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 7, 31}));
  BOOST_TEST(767224 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 8, 1}));
  BOOST_TEST(767254 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 8, 31}));
  BOOST_TEST(767255 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 9, 1}));
  BOOST_TEST(767284 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 9, 30}));
  BOOST_TEST(767285 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 10, 1}));
  BOOST_TEST(767315 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 10, 31}));
  BOOST_TEST(767316 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 11, 1}));
  BOOST_TEST(767345 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 11, 30}));
  BOOST_TEST(767346 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 12, 1}));
  BOOST_TEST(767376 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4000, 12, 31}));
  BOOST_TEST(884254 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 1, 1}));
  BOOST_TEST(884284 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 1, 31}));
  BOOST_TEST(884285 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 2, 1}));
  BOOST_TEST(884312 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 2, 28}));
  BOOST_TEST(884313 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 3, 1}));
  BOOST_TEST(884343 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 3, 31}));
  BOOST_TEST(884344 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 4, 1}));
  BOOST_TEST(884373 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 4, 30}));
  BOOST_TEST(884374 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 5, 1}));
  BOOST_TEST(884404 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 5, 31}));
  BOOST_TEST(884405 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 6, 1}));
  BOOST_TEST(884434 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 6, 30}));
  BOOST_TEST(884435 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 7, 1}));
  BOOST_TEST(884465 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 7, 31}));
  BOOST_TEST(884466 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 8, 1}));
  BOOST_TEST(884496 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 8, 31}));
  BOOST_TEST(884497 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 9, 1}));
  BOOST_TEST(884526 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 9, 30}));
  BOOST_TEST(884527 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 10, 1}));
  BOOST_TEST(884557 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 10, 31}));
  BOOST_TEST(884558 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 11, 1}));
  BOOST_TEST(884587 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 11, 30}));
  BOOST_TEST(884588 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 12, 1}));
  BOOST_TEST(884618 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{4321, 12, 31}));
  BOOST_TEST(2958101 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 1, 1}));
  BOOST_TEST(2958131 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 1, 31}));
  BOOST_TEST(2958132 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 2, 1}));
  BOOST_TEST(2958159 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 2, 28}));
  BOOST_TEST(2958160 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 3, 1}));
  BOOST_TEST(2958190 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 3, 31}));
  BOOST_TEST(2958191 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 4, 1}));
  BOOST_TEST(2958220 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 4, 30}));
  BOOST_TEST(2958221 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 5, 1}));
  BOOST_TEST(2958251 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 5, 31}));
  BOOST_TEST(2958252 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 6, 1}));
  BOOST_TEST(2958281 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 6, 30}));
  BOOST_TEST(2958282 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 7, 1}));
  BOOST_TEST(2958312 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 7, 31}));
  BOOST_TEST(2958313 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 8, 1}));
  BOOST_TEST(2958343 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 8, 31}));
  BOOST_TEST(2958344 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 9, 1}));
  BOOST_TEST(2958373 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 9, 30}));
  BOOST_TEST(2958374 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 10, 1}));
  BOOST_TEST(2958404 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 10, 31}));
  BOOST_TEST(2958405 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 11, 1}));
  BOOST_TEST(2958434 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 11, 30}));
  BOOST_TEST(2958435 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 12, 1}));
  BOOST_TEST(2958465 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{9999, 12, 31}));
}

BOOST_AUTO_TEST_CASE(datetime_date_only_1904, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 1, 1}, true));
  BOOST_TEST(30 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 1, 31}, true));
  BOOST_TEST(31 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 2, 1}, true));
  BOOST_TEST(59 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 2, 29}, true));
  BOOST_TEST(60 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 3, 1}, true));
  BOOST_TEST(90 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 3, 31}, true));
  BOOST_TEST(91 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 4, 1}, true));
  BOOST_TEST(120 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 4, 30}, true));
  BOOST_TEST(121 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 5, 1}, true));
  BOOST_TEST(151 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 5, 31}, true));
  BOOST_TEST(152 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 6, 1}, true));
  BOOST_TEST(181 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 6, 30}, true));
  BOOST_TEST(182 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 7, 1}, true));
  BOOST_TEST(212 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 7, 31}, true));
  BOOST_TEST(213 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 8, 1}, true));
  BOOST_TEST(243 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 8, 31}, true));
  BOOST_TEST(244 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 9, 1}, true));
  BOOST_TEST(273 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 9, 30}, true));
  BOOST_TEST(274 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 10, 1}, true));
  BOOST_TEST(304 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 10, 31}, true));
  BOOST_TEST(305 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 11, 1}, true));
  BOOST_TEST(334 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 11, 30}, true));
  BOOST_TEST(335 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 12, 1}, true));
  BOOST_TEST(365 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1904, 12, 31}, true));
  BOOST_TEST(1153 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1907, 2, 27}, true));
  BOOST_TEST(1154 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1907, 2, 28}, true));
  BOOST_TEST(1155 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1907, 3, 1}, true));
  BOOST_TEST(1156 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1907, 3, 2}, true));
  BOOST_TEST(1157 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1907, 3, 3}, true));
  BOOST_TEST(1158 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1907, 3, 4}, true));
  BOOST_TEST(1159 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1907, 3, 5}, true));
  BOOST_TEST(1160 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1907, 3, 6}, true));
  BOOST_TEST(34699 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 1, 1}, true));
  BOOST_TEST(34729 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 1, 31}, true));
  BOOST_TEST(34730 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 2, 1}, true));
  BOOST_TEST(34757 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 2, 28}, true));
  BOOST_TEST(34758 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 3, 1}, true));
  BOOST_TEST(34788 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 3, 31}, true));
  BOOST_TEST(34789 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 4, 1}, true));
  BOOST_TEST(34818 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 4, 30}, true));
  BOOST_TEST(34819 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 5, 1}, true));
  BOOST_TEST(34849 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 5, 31}, true));
  BOOST_TEST(34850 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 6, 1}, true));
  BOOST_TEST(34879 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 6, 30}, true));
  BOOST_TEST(34880 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 7, 1}, true));
  BOOST_TEST(34910 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 7, 31}, true));
  BOOST_TEST(34911 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 8, 1}, true));
  BOOST_TEST(34941 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 8, 31}, true));
  BOOST_TEST(34942 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 9, 1}, true));
  BOOST_TEST(34971 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 9, 30}, true));
  BOOST_TEST(34972 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 10, 1}, true));
  BOOST_TEST(35002 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 10, 31}, true));
  BOOST_TEST(35003 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 11, 1}, true));
  BOOST_TEST(35032 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 11, 30}, true));
  BOOST_TEST(35033 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 12, 1}, true));
  BOOST_TEST(35063 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{1999, 12, 31}, true));
  BOOST_TEST(35064 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 1, 1}, true));
  BOOST_TEST(35094 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 1, 31}, true));
  BOOST_TEST(35095 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 2, 1}, true));
  BOOST_TEST(35123 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 2, 29}, true));
  BOOST_TEST(35124 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 3, 1}, true));
  BOOST_TEST(35154 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 3, 31}, true));
  BOOST_TEST(35155 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 4, 1}, true));
  BOOST_TEST(35184 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 4, 30}, true));
  BOOST_TEST(35185 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 5, 1}, true));
  BOOST_TEST(35215 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 5, 31}, true));
  BOOST_TEST(35216 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 6, 1}, true));
  BOOST_TEST(35245 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 6, 30}, true));
  BOOST_TEST(35246 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 7, 1}, true));
  BOOST_TEST(35276 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 7, 31}, true));
  BOOST_TEST(35277 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 8, 1}, true));
  BOOST_TEST(35307 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 8, 31}, true));
  BOOST_TEST(35308 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 9, 1}, true));
  BOOST_TEST(35337 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 9, 30}, true));
  BOOST_TEST(35338 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 10, 1}, true));
  BOOST_TEST(35368 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 10, 31}, true));
  BOOST_TEST(35369 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 11, 1}, true));
  BOOST_TEST(35398 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 11, 30}, true));
  BOOST_TEST(35399 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 12, 1}, true));
  BOOST_TEST(35429 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2000, 12, 31}, true));
  BOOST_TEST(35430 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 1, 1}, true));
  BOOST_TEST(35460 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 1, 31}, true));
  BOOST_TEST(35461 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 2, 1}, true));
  BOOST_TEST(35488 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 2, 28}, true));
  BOOST_TEST(35489 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 3, 1}, true));
  BOOST_TEST(35519 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 3, 31}, true));
  BOOST_TEST(35520 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 4, 1}, true));
  BOOST_TEST(35549 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 4, 30}, true));
  BOOST_TEST(35550 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 5, 1}, true));
  BOOST_TEST(35580 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 5, 31}, true));
  BOOST_TEST(35581 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 6, 1}, true));
  BOOST_TEST(35610 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 6, 30}, true));
  BOOST_TEST(35611 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 7, 1}, true));
  BOOST_TEST(35641 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 7, 31}, true));
  BOOST_TEST(35642 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 8, 1}, true));
  BOOST_TEST(35672 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 8, 31}, true));
  BOOST_TEST(35673 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 9, 1}, true));
  BOOST_TEST(35702 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 9, 30}, true));
  BOOST_TEST(35703 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 10, 1}, true));
  BOOST_TEST(35733 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 10, 31}, true));
  BOOST_TEST(35734 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 11, 1}, true));
  BOOST_TEST(35763 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 11, 30}, true));
  BOOST_TEST(35764 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 12, 1}, true));
  BOOST_TEST(35794 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2001, 12, 31}, true));
  BOOST_TEST(181161 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 1, 1}, true));
  BOOST_TEST(181191 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 1, 31}, true));
  BOOST_TEST(181192 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 2, 1}, true));
  BOOST_TEST(181220 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 2, 29}, true));
  BOOST_TEST(181221 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 3, 1}, true));
  BOOST_TEST(181251 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 3, 31}, true));
  BOOST_TEST(181252 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 4, 1}, true));
  BOOST_TEST(181281 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 4, 30}, true));
  BOOST_TEST(181282 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 5, 1}, true));
  BOOST_TEST(181312 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 5, 31}, true));
  BOOST_TEST(181313 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 6, 1}, true));
  BOOST_TEST(181342 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 6, 30}, true));
  BOOST_TEST(181343 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 7, 1}, true));
  BOOST_TEST(181373 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 7, 31}, true));
  BOOST_TEST(181374 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 8, 1}, true));
  BOOST_TEST(181404 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 8, 31}, true));
  BOOST_TEST(181405 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 9, 1}, true));
  BOOST_TEST(181434 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 9, 30}, true));
  BOOST_TEST(181435 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 10, 1}, true));
  BOOST_TEST(181465 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 10, 31}, true));
  BOOST_TEST(181466 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 11, 1}, true));
  BOOST_TEST(181495 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 11, 30}, true));
  BOOST_TEST(181496 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 12, 1}, true));
  BOOST_TEST(181526 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{2400, 12, 31}, true));
  BOOST_TEST(765549 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 1, 1}, true));
  BOOST_TEST(765579 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 1, 31}, true));
  BOOST_TEST(765580 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 2, 1}, true));
  BOOST_TEST(765608 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 2, 29}, true));
  BOOST_TEST(765609 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 3, 1}, true));
  BOOST_TEST(765639 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 3, 31}, true));
  BOOST_TEST(765640 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 4, 1}, true));
  BOOST_TEST(765669 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 4, 30}, true));
  BOOST_TEST(765670 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 5, 1}, true));
  BOOST_TEST(765700 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 5, 31}, true));
  BOOST_TEST(765701 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 6, 1}, true));
  BOOST_TEST(765730 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 6, 30}, true));
  BOOST_TEST(765731 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 7, 1}, true));
  BOOST_TEST(765761 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 7, 31}, true));
  BOOST_TEST(765762 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 8, 1}, true));
  BOOST_TEST(765792 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 8, 31}, true));
  BOOST_TEST(765793 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 9, 1}, true));
  BOOST_TEST(765822 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 9, 30}, true));
  BOOST_TEST(765823 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 10, 1}, true));
  BOOST_TEST(765853 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 10, 31}, true));
  BOOST_TEST(765854 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 11, 1}, true));
  BOOST_TEST(765883 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 11, 30}, true));
  BOOST_TEST(765884 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 12, 1}, true));
  BOOST_TEST(765914 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4000, 12, 31}, true));
  BOOST_TEST(882792 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 1, 1}, true));
  BOOST_TEST(882822 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 1, 31}, true));
  BOOST_TEST(882823 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 2, 1}, true));
  BOOST_TEST(882850 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 2, 28}, true));
  BOOST_TEST(882851 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 3, 1}, true));
  BOOST_TEST(882881 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 3, 31}, true));
  BOOST_TEST(882882 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 4, 1}, true));
  BOOST_TEST(882911 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 4, 30}, true));
  BOOST_TEST(882912 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 5, 1}, true));
  BOOST_TEST(882942 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 5, 31}, true));
  BOOST_TEST(882943 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 6, 1}, true));
  BOOST_TEST(882972 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 6, 30}, true));
  BOOST_TEST(882973 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 7, 1}, true));
  BOOST_TEST(883003 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 7, 31}, true));
  BOOST_TEST(883004 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 8, 1}, true));
  BOOST_TEST(883034 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 8, 31}, true));
  BOOST_TEST(883035 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 9, 1}, true));
  BOOST_TEST(883064 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 9, 30}, true));
  BOOST_TEST(883065 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 10, 1}, true));
  BOOST_TEST(883095 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 10, 31}, true));
  BOOST_TEST(883096 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 11, 1}, true));
  BOOST_TEST(883125 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 11, 30}, true));
  BOOST_TEST(883126 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 12, 1}, true));
  BOOST_TEST(883156 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{4321, 12, 31}, true));
  BOOST_TEST(2956639 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 1, 1}, true));
  BOOST_TEST(2956669 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 1, 31}, true));
  BOOST_TEST(2956670 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 2, 1}, true));
  BOOST_TEST(2956697 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 2, 28}, true));
  BOOST_TEST(2956698 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 3, 1}, true));
  BOOST_TEST(2956728 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 3, 31}, true));
  BOOST_TEST(2956729 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 4, 1}, true));
  BOOST_TEST(2956758 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 4, 30}, true));
  BOOST_TEST(2956759 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 5, 1}, true));
  BOOST_TEST(2956789 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 5, 31}, true));
  BOOST_TEST(2956790 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 6, 1}, true));
  BOOST_TEST(2956819 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 6, 30}, true));
  BOOST_TEST(2956820 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 7, 1}, true));
  BOOST_TEST(2956850 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 7, 31}, true));
  BOOST_TEST(2956851 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 8, 1}, true));
  BOOST_TEST(2956881 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 8, 31}, true));
  BOOST_TEST(2956882 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 9, 1}, true));
  BOOST_TEST(2956911 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 9, 30}, true));
  BOOST_TEST(2956912 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 10, 1}, true));
  BOOST_TEST(2956942 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 10, 31}, true));
  BOOST_TEST(2956943 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 11, 1}, true));
  BOOST_TEST(2956972 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 11, 30}, true));
  BOOST_TEST(2956973 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 12, 1}, true));
  BOOST_TEST(2957003 == xwpp::datetime_to_excel_date_with_epoch(xwpp::datetime_t{9999, 12, 31}, true));
}

BOOST_AUTO_TEST_CASE(datetime_time_only, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 0, 0, 0.000}));
  BOOST_TEST(1.0650613425925924E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 0, 15, 20.213}));
  BOOST_TEST(1.1670023148148148E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 0, 16, 48.290}));
  BOOST_TEST(3.8488958333333337E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 0, 55, 25.446}));
  BOOST_TEST(4.3598275462962965E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 1, 2, 46.891}));
  BOOST_TEST(4.4624965277777782E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 1, 4, 15.597}));
  BOOST_TEST(4.8389918981481483E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 1, 9, 40.889}));
  BOOST_TEST(4.9682407407407404E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 1, 11, 32.560}));
  BOOST_TEST(6.2721863425925936E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 1, 30, 19.169}));
  BOOST_TEST(7.5296064814814809E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 1, 48, 25.580}));
  BOOST_TEST(8.5786099537037031E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 2, 3, 31.919}));
  BOOST_TEST(9.1110949074074077E-2 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 2, 11, 11.986}));
  BOOST_TEST(0.10042934027777778 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 2, 24, 37.095}));
  BOOST_TEST(0.1077224537037037 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 2, 35, 7.220}));
  BOOST_TEST(0.11472348379629631 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 2, 45, 12.109}));
  BOOST_TEST(0.12962951388888888 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 3, 6, 39.990}));
  BOOST_TEST(0.13065105324074075 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 3, 8, 8.251}));
  BOOST_TEST(0.13833999999999999 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 3, 19, 12.576}));
  BOOST_TEST(0.14563164351851851 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 3, 29, 42.574}));
  BOOST_TEST(0.1510510763888889 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 3, 37, 30.813}));
  BOOST_TEST(0.1768313773148148 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 4, 14, 38.231}));
  BOOST_TEST(0.17810832175925925 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 4, 16, 28.559}));
  BOOST_TEST(0.17914608796296297 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 4, 17, 58.222}));
  BOOST_TEST(0.18173372685185185 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 4, 21, 41.794}));
  BOOST_TEST(0.2059698148148148 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 4, 56, 35.792}));
  BOOST_TEST(0.22586672453703704 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 5, 25, 14.885}));
  BOOST_TEST(0.22645513888888891 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 5, 26, 5.724}));
  BOOST_TEST(0.24078782407407406 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 5, 46, 44.068}));
  BOOST_TEST(0.2416798726851852 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 5, 48, 1.141}));
  BOOST_TEST(0.24574438657407408 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 5, 53, 52.315}));
  BOOST_TEST(0.26028449074074073 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 6, 14, 48.580}));
  BOOST_TEST(0.28212659722222222 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 6, 46, 15.738}));
  BOOST_TEST(0.31343063657407405 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 7, 31, 20.407}));
  BOOST_TEST(0.33233511574074076 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 7, 58, 33.754}));
  BOOST_TEST(0.33869363425925925 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 8, 7, 43.130}));
  BOOST_TEST(0.35360059027777774 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 8, 29, 11.091}));
  BOOST_TEST(0.380732962962963 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 9, 8, 15.328}));
  BOOST_TEST(0.39631690972222228 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 9, 30, 41.781}));
  BOOST_TEST(0.39866275462962958 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 9, 34, 4.462}));
  BOOST_TEST(0.40097158564814817 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 9, 37, 23.945}));
  BOOST_TEST(0.40135017361111114 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 9, 37, 56.655}));
  BOOST_TEST(0.40639155092592594 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 9, 45, 12.230}));
  BOOST_TEST(0.41267108796296298 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 9, 54, 14.782}));
  BOOST_TEST(0.41275587962962962 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 9, 54, 22.108}));
  BOOST_TEST(0.41777952546296299 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 10, 1, 36.151}));
  BOOST_TEST(0.50681252314814818 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 12, 9, 48.602}));
  BOOST_TEST(0.52371005787037039 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 12, 34, 8.549}));
  BOOST_TEST(0.53896406249999995 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 12, 56, 6.495}));
  BOOST_TEST(0.54095158564814816 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 12, 58, 58.217}));
  BOOST_TEST(0.54160026620370372 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 12, 59, 54.263}));
  BOOST_TEST(0.56575614583333333 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 13, 34, 41.331}));
  BOOST_TEST(0.58227547453703699 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 13, 58, 28.601}));
  BOOST_TEST(0.58491781249999997 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 14, 2, 16.899}));
  BOOST_TEST(0.60853523148148148 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 14, 36, 17.444}));
  BOOST_TEST(0.60969271990740748 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 14, 37, 57.451}));
  BOOST_TEST(0.6234115393518519 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 14, 57, 42.757}));
  BOOST_TEST(0.6325035532407407 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 15, 10, 48.307}));
  BOOST_TEST(0.63518391203703706 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 15, 14, 39.890}));
  BOOST_TEST(0.63874986111111109 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 15, 19, 47.988}));
  BOOST_TEST(0.66972620370370362 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 16, 4, 24.344}));
  BOOST_TEST(0.68222166666666662 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 16, 22, 23.952}));
  BOOST_TEST(0.6874536921296297 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 16, 29, 55.999}));
  BOOST_TEST(0.70717892361111112 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 16, 58, 20.259}));
  BOOST_TEST(0.71113906250000003 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 17, 4, 2.415}));
  BOOST_TEST(0.72117627314814825 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 17, 18, 29.630}));
  BOOST_TEST(0.74121901620370367 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 17, 47, 21.323}));
  BOOST_TEST(0.74548456018518516 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 17, 53, 29.866}));
  BOOST_TEST(0.74561430555555563 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 17, 53, 41.076}));
  BOOST_TEST(0.74659773148148145 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 17, 55, 6.044}));
  BOOST_TEST(0.760291099537037 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 18, 14, 49.151}));
  BOOST_TEST(0.76233493055555546 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 18, 17, 45.738}));
  BOOST_TEST(0.77082986111111118 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 18, 29, 59.700}));
  BOOST_TEST(0.77316241898148153 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 18, 33, 21.233}));
  BOOST_TEST(0.80167445601851861 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 19, 14, 24.673}));
  BOOST_TEST(0.80362055555555545 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 19, 17, 12.816}));
  BOOST_TEST(0.80806039351851855 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 19, 23, 36.418}));
  BOOST_TEST(0.82391097222222232 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 19, 46, 25.908}));
  BOOST_TEST(0.83874206018518516 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 20, 7, 47.314}));
  BOOST_TEST(0.85529633101851854 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 20, 31, 37.603}));
  BOOST_TEST(0.86108530092592594 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 20, 39, 57.770}));
  BOOST_TEST(0.86825309027777775 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 20, 50, 17.067}));
  BOOST_TEST(0.87705818287037041 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 21, 2, 57.827}));
  BOOST_TEST(0.891036099537037 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 21, 23, 5.519}));
  BOOST_TEST(0.89918486111111118 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 21, 34, 49.572}));
  BOOST_TEST(0.90215212962962965 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 21, 39, 5.944}));
  BOOST_TEST(0.90229659722222222 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 21, 39, 18.426}));
  BOOST_TEST(0.90703436342592603 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 21, 46, 7.769}));
  BOOST_TEST(0.91522756944444439 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 21, 57, 55.662}));
  BOOST_TEST(0.92999689814814823 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 22, 19, 11.732}));
  BOOST_TEST(0.93323351851851843 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 22, 23, 51.376}));
  BOOST_TEST(0.93609688657407408 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 22, 27, 58.771}));
  BOOST_TEST(0.94687953703703709 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 22, 43, 30.392}));
  BOOST_TEST(0.95029900462962968 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 22, 48, 25.834}));
  BOOST_TEST(0.95407091435185187 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 22, 53, 51.727}));
  BOOST_TEST(0.96732101851851848 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 23, 12, 56.536}));
  BOOST_TEST(0.96937626157407408 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 23, 15, 54.109}));
  BOOST_TEST(0.97028509259259266 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 23, 17, 12.632}));
  BOOST_TEST(0.99999998842592586 == xwpp::datetime_to_excel_datetime(xwpp::datetime_t{0, 0, 0, 23, 59, 59.999}));
}

BOOST_AUTO_TEST_CASE(unixtime, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::unixtime_to_excel_date(-2209075200LL));        // 1899-12-31 00:00
  BOOST_TEST(0.5 == xwpp::unixtime_to_excel_date(-2209032000LL));      // 1899-12-31 12:00
  BOOST_TEST(1 == xwpp::unixtime_to_excel_date(-2208988800LL));        // 1900-01-01 00:00
  BOOST_TEST(1.5 == xwpp::unixtime_to_excel_date(-2208945600LL));      // 1900-01-01 12:00
  BOOST_TEST(59 == xwpp::unixtime_to_excel_date(-2203977600LL));       // 1900-02-28 00:00
  BOOST_TEST(59.5 == xwpp::unixtime_to_excel_date(-2203934400LL));     // 1900-02-28 12:00
  BOOST_TEST(61 == xwpp::unixtime_to_excel_date(-2203891200LL));       // 1900-03-01 00:00
  BOOST_TEST(61.5 == xwpp::unixtime_to_excel_date(-2203848000LL));     // 1900-03-01 12:00
  BOOST_TEST(25569 == xwpp::unixtime_to_excel_date(0));                // 1970-01-01 00:00
  BOOST_TEST(25569.5 == xwpp::unixtime_to_excel_date(43200));          // 1970-01-01 12:00
  BOOST_TEST(36526 == xwpp::unixtime_to_excel_date(946684800));        // 2000-01-01 00:00
  BOOST_TEST(36526.5 == xwpp::unixtime_to_excel_date(946728000));      // 2000-01-01 12:00
  BOOST_TEST(2958101 == xwpp::unixtime_to_excel_date(253370764800));   // 9999-01-01 00:00
  BOOST_TEST(2958101.5 == xwpp::unixtime_to_excel_date(253370808000)); // 9999-01-01 12:00
}

BOOST_AUTO_TEST_CASE(unixtime_1904, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::unixtime_to_excel_date_with_epoch(-2082844800, true));     // 1904-01-01 00:00
  BOOST_TEST(0.5 == xwpp::unixtime_to_excel_date_with_epoch(-2082801600, true));   // 1904-01-01 12:00
  BOOST_TEST(24107 == xwpp::unixtime_to_excel_date_with_epoch(0, true));           // 1970-01-01 00:00
  BOOST_TEST(24107.5 == xwpp::unixtime_to_excel_date_with_epoch(43200, true));     // 1970-01-01 12:00
  BOOST_TEST(35064 == xwpp::unixtime_to_excel_date_with_epoch(946684800, true));   // 2000-01-01 00:00
  BOOST_TEST(35064.5 == xwpp::unixtime_to_excel_date_with_epoch(946728000, true)); // 2000-01-01 12:00
}

BOOST_AUTO_TEST_SUITE_END()
