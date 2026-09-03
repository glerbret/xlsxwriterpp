/*
 * chartsheet - A Xlsxwriter++ library for creating Excel XLSX chartsheet files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @page sheet_page The objects used by worksheets and chartsheets.
 *
 * @see @ref sheet.h for full details of the functionality.
 *
 * @file
 *
 * @brief Common object to `worksheet_t` and `chartsheet_t`.
 *
 * Common object to `worksheet_t` and `chartsheet_t`.
 */

#ifndef XWPP_SHEET_H
#define XWPP_SHEET_H

#include "xwpp/chart.h"
#include "xwpp/common.h"
#include "xwpp/drawing.h"
#include "xwpp/format.h"
#include "xwpp/shared_strings.h"

#include <array>
#include <cstddef>
#include <list>
#include <map>
#include <optional>
#include <string>
#include <tuple>
#include <vector>

/// @cond
namespace xwpp
{
/// @endcond

/** Default Excel column width in character units. */
const double DEF_COL_WIDTH{8.43};

/** Default Excel row height in character units. */
const double DEF_ROW_HEIGHT{15.0};

/**
 * @brief Options for headers and footers.
 *
 * Optional parameters used in the `worksheet_t::set_header()`,
 * `worksheet_t::set_footer()`, `chartsheet_t::set_header()` and
 * `chartsheet_t::set_footer()` functions.
 */
struct header_footer_options_t
{
  /**
   * Header or footer margin in inches. Excel default is 0.3. Must by
   * larger than 0.0.
   *
   * @see `worksheet_t::set_header()`.
   */
  double margin_{0.};

  /**
   * The left header image filename, with path if required. This should
   * have a corresponding `&G` placeholder in the `&L` section of
   * the header/footer string.
   *
   * @see `worksheet_t::set_header()`.
   */
  std::string image_left_;

  /**
   * The center header image filename, with path if required. This should
   * have a corresponding `&G` placeholder in the `&C` section of
   * the header/footer string.
   *
   * @see `worksheet_t::set_header()`.
   */
  std::string image_center_;

  /**
   * The right header image filename, with path if required. This should
   * have a corresponding `&G` placeholder in the `&R` section of
   * the header/footer string.
   *
   * @see `worksheet_t::set_header()`.
   */
  std::string image_right_;
};

/**
 * @brief Options to control the positioning of objects.
 *
 * Options to control the positioning of sheet objects such as images
 * or charts. See @ref working_with_object_positioning.
 */
enum class object_position_t
{
  /** Default positioning for the object. */
  DEFAULT,

  /** Move and size the worksheet object with the cells. */
  MOVE_AND_SIZE,

  /** Move but don't size the worksheet object with the cells. */
  MOVE_DONT_SIZE,

  /** Don't move or size the worksheet object with the cells. */
  DONT_MOVE_DONT_SIZE,

  /** Same as `MOVE_AND_SIZE` except `Xlsxwriter++` applies hidden
   *  cells after the object is inserted. */
  MOVE_AND_SIZE_AFTER
};

/**
 * @brief Display type for a cell comment.
 *
 * Set the display type for a cell comment. This is hidden by default but
 * can be set to visible with the `worksheet_t::show_comments()` function.
 */
enum class comment_display_t
{
  /** Default to the worksheet default which can be hidden or visible. */
  DEFAULT,

  /**
   * Hide the cell comment.
   *
   * Usually the default.
   */
  HIDDEN,

  /**
   * Show the cell comment.
   *
   * Can also be set for the worksheet with the `worksheet_t::show_comments()` function.
   */
  VISIBLE
};

/**
 * @brief Worksheet protection options.
 *
 * Worksheet protection options.
 */
struct protection_t
{
  /**
   * Turn off selection of locked cells. This in on in Excel by default.
   */
  bool no_select_locked_cells_{false};

  /**
   * Turn off selection of unlocked cells. This in on in Excel by default.
   */
  bool no_select_unlocked_cells_{false};

  /**
   * Prevent formatting of cells.
   */
  bool format_cells_{false};

  /**
   * Prevent formatting of columns.
   */
  bool format_columns_{false};

  /**
   * Prevent formatting of rows.
   */
  bool format_rows_{false};

  /**
   * Prevent insertion of columns.
   */
  bool insert_columns_{false};

  /**
   * Prevent insertion of rows.
   */
  bool insert_rows_{false};

  /**
   * Prevent insertion of hyperlinks.
   */
  bool insert_hyperlinks_{false};

  /**
   * Prevent deletion of columns.
   */
  bool delete_columns_{false};

  /**
   * Prevent deletion of rows.
   */
  bool delete_rows_{false};

  /**
   * Prevent sorting data.
   */
  bool sort_{false};

  /**
   * Prevent filtering data.
   */
  bool autofilter_{false};

  /**
   * Prevent insertion of pivot tables.
   */
  bool pivot_tables_{false};

  /**
   * Protect scenarios.
   */
  bool scenarios_{false};

  /**
   * Protect drawing objects. Worksheets only.
   *
   * @todo Not clear, why there is two options `%objects_` and `%no_objects_`.
   * @todo And `%objects_` seems to be wrong name as XML object option is set to 1 iff objects_ is false.
   */
  bool objects_{false};

  /**
   * Turn off chartsheet content protection.
   */
  bool no_content_{false};

  /**
   * Turn off chartsheet objects.
   */
  bool no_objects_{false};
};

enum class image_position_t
{
  HEADER_LEFT = 0,
  HEADER_CENTER,
  HEADER_RIGHT,
  FOOTER_LEFT,
  FOOTER_CENTER,
  FOOTER_RIGHT
};

// Worksheet initialization data.
struct sheet_init_data_t
{
  size_t index_{0};
  bool hidden_{false};
  size_t* active_sheet_{nullptr};
  size_t* first_sheet_{nullptr};
  shared_strings_t* sst_{nullptr};
  std::string name_;
  std::string quoted_name_;
  format_t* default_url_format_{nullptr};
  uint16_t max_url_length_{2079};
  bool use_1904_epoch_{false};
};

// Internal struct to copy protection options and internal metadata.
struct protection_obj_t
{
  bool no_select_locked_cells_{false};
  bool no_select_unlocked_cells_{false};
  bool format_cells_{false};
  bool format_columns_{false};
  bool format_rows_{false};
  bool insert_columns_{false};
  bool insert_rows_{false};
  bool insert_hyperlinks_{false};
  bool delete_columns_{false};
  bool delete_rows_{false};
  bool sort_{false};
  bool autofilter_{false};
  bool pivot_tables_{false};
  bool scenarios_{false};
  bool objects_{false};
  bool no_content_{false};
  bool no_objects_{false};
  bool no_sheet_{false};
  bool is_configured_{false};
  std::string hash_;
};

// Internal struct to represent `image_options_t` and `chart_options_t`
// values as well as internal metadata.
struct object_properties_t
{
  int32_t x_offset_{0};
  int32_t y_offset_{0};
  double x_scale_{1.};
  double y_scale_{1.};
  row_num_t row_num_{0};
  col_num_t col_num_{0};
  std::string filename_;
  std::string description_;
  std::string url_;
  std::string tip_;
  object_position_t object_position_{object_position_t::DEFAULT};
  image_types_t image_type_{image_types_t::UNKNOWN};
  std::vector<unsigned char> image_buffer_;
  double width_{0.};
  double height_{0.};
  std::string extension_;
  double x_dpi_{0.};
  double y_dpi_{0.};
  chart_t* chart_{nullptr};
  bool is_duplicate_{false};
  bool is_background_{false};
  std::string md5_;
  std::string image_position_;
  bool decorative_{false};
  format_t* format_{nullptr};
};

// Internal structure for VML object options.
struct vml_obj_t
{
  row_num_t row_num_{0};
  col_num_t col_num_{0};
  row_num_t start_row_{0};
  col_num_t start_col_{0};
  int32_t x_offset_{0};
  int32_t y_offset_{0};
  uint64_t col_absolute_{0};
  uint64_t row_absolute_{0};
  uint32_t width_{0};
  uint32_t height_{0};
  double x_dpi_{0.};
  double y_dpi_{0.};
  color_t color_{color_t::UNSET};
  uint8_t font_family_{0};
  comment_display_t visible_{comment_display_t::DEFAULT};
  uint32_t author_id_{0};
  uint32_t rel_index_{0};
  double font_size_{0.};
  drawing_coords_t from_;
  drawing_coords_t to_;
  std::string author_;
  std::string font_name_;
  std::string text_;
  std::string image_position_;
  std::string name_;
  std::string macro_;
};

class sheet_t
{
public:
  virtual ~sheet_t() = default;

  /**
   * @brief Set the color of the sheet tab.
   *
   * @param color The tab color.
   *
   * The `%set_tab_color()` function is used to change the color of
   * the sheet tab:
   *
   * @code
   *  worksheet1.set_tab_color(xwpp::color_t::RED);
   *  worksheet2.set_tab_color(xwpp::color_t::GREEN);
   *  worksheet3.set_tab_color(xwpp::color_t(0xFF9900)); // Orange.
   * @endcode
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   */
  void set_tab_color(color_t color);

  /**
   * @brief Set the sheet zoom factor.
   *
   * @param scale zoom factor.
   *
   * Set the sheet zoom factor in the range `10 <= zoom <= 400`:
   *
   * @code
   *  worksheet1.set_zoom(50);
   *  worksheet2.set_zoom(75);
   *  worksheet3.set_zoom(300);
   *  worksheet4.set_zoom(400);
   * @endcode
   *
   * The default zoom factor is 100. It isn't possible to set the zoom to
   * "Selection" because it is calculated by Excel at run-time.
   *
   * @note `%set_zoom()` does not affect the scale of the printed
   * page. For that you should use `set_print_scale()`.
   */
  void set_zoom(uint16_t scale);

  /**
   * @brief Set the paper type for printing.
   *
   * @param paper_size The Excel paper format type.
   *
   * This function is used to set the paper format for the printed output of a
   * sheet. The following paper styles are available:
   *
   *   Index    | Paper format            | Paper size
   *   :------- | :---------------------- | :-------------------
   *   0        | Printer default         | Printer default
   *   1        | Letter                  | 8 1/2 x 11 in
   *   2        | Letter Small            | 8 1/2 x 11 in
   *   3        | Tabloid                 | 11 x 17 in
   *   4        | Ledger                  | 17 x 11 in
   *   5        | Legal                   | 8 1/2 x 14 in
   *   6        | Statement               | 5 1/2 x 8 1/2 in
   *   7        | Executive               | 7 1/4 x 10 1/2 in
   *   8        | A3                      | 297 x 420 mm
   *   9        | A4                      | 210 x 297 mm
   *   10       | A4 Small                | 210 x 297 mm
   *   11       | A5                      | 148 x 210 mm
   *   12       | B4                      | 250 x 354 mm
   *   13       | B5                      | 182 x 257 mm
   *   14       | Folio                   | 8 1/2 x 13 in
   *   15       | Quarto                  | 215 x 275 mm
   *   16       | ---                     | 10x14 in
   *   17       | ---                     | 11x17 in
   *   18       | Note                    | 8 1/2 x 11 in
   *   19       | Envelope 9              | 3 7/8 x 8 7/8
   *   20       | Envelope 10             | 4 1/8 x 9 1/2
   *   21       | Envelope 11             | 4 1/2 x 10 3/8
   *   22       | Envelope 12             | 4 3/4 x 11
   *   23       | Envelope 14             | 5 x 11 1/2
   *   24       | C size sheet            | ---
   *   25       | D size sheet            | ---
   *   26       | E size sheet            | ---
   *   27       | Envelope DL             | 110 x 220 mm
   *   28       | Envelope C3             | 324 x 458 mm
   *   29       | Envelope C4             | 229 x 324 mm
   *   30       | Envelope C5             | 162 x 229 mm
   *   31       | Envelope C6             | 114 x 162 mm
   *   32       | Envelope C65            | 114 x 229 mm
   *   33       | Envelope B4             | 250 x 353 mm
   *   34       | Envelope B5             | 176 x 250 mm
   *   35       | Envelope B6             | 176 x 125 mm
   *   36       | Envelope                | 110 x 230 mm
   *   37       | Monarch                 | 3.875 x 7.5 in
   *   38       | Envelope                | 3 5/8 x 6 1/2 in
   *   39       | Fanfold                 | 14 7/8 x 11 in
   *   40       | German Std Fanfold      | 8 1/2 x 12 in
   *   41       | German Legal Fanfold    | 8 1/2 x 13 in
   *
   * @code
   *  worksheet1.set_paper(1);  // US Letter
   *  worksheet2.set_paper(9);  // A4
   * @endcode
   *
   * If you do not specify a paper type the sheet will print using the
   * printer's default paper style.
   *
   * @note It is likely that not all of these paper types will be available to
   * the end user since it will depend on the paper formats that the user's
   * printer supports. Therefore, it is best to stick to standard paper types:
   *
   * @todo Use enum as `paper_size`.
   */
  void set_paper(uint8_t paper_size);

  /**
   * @brief Set the printed page header caption.
   *
   * @param str     The header string.
   * @param options Header options.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * Headers and footers are generated using a string which is a combination of
   * plain text and control characters.
   *
   * The available control character are:
   *
   *   | Control         | Category      | Description           |
   *   | --------------- | ------------- | --------------------- |
   *   | `&L`            | Justification | Left                  |
   *   | `&C`            |               | Center                |
   *   | `&R`            |               | Right                 |
   *   | `&P`            | Information   | Page number           |
   *   | `&N`            |               | Total number of pages |
   *   | `&D`            |               | Date                  |
   *   | `&T`            |               | Time                  |
   *   | `&F`            |               | File name             |
   *   | `&A`            |               | Worksheet name        |
   *   | `&Z`            |               | Workbook path         |
   *   | `&fontsize`     | Font          | Font size             |
   *   | `&"font,style"` |               | Font name and style   |
   *   | `&U`            |               | Single underline      |
   *   | `&E`            |               | Double underline      |
   *   | `&S`            |               | Strikethrough         |
   *   | `&X`            |               | Superscript           |
   *   | `&Y`            |               | Subscript             |
   *   | `&G`            | Images        | Image placeholder     |
   *   | `&&`            | Miscellaneous | Literal ampersand &   |
   *
   * Text in headers and footers can be justified (aligned) to the left, center
   * and right by prefixing the text with the control characters `&L`, `&C` and
   * `&R`.
   *
   * For example (with ASCII art representation of the results):
   *
   * @code
   *  worksheet.set_header("&LHello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    | Hello                                                         |
   *  //    |                                                               |
   *
   *
   *  worksheet.set_header("&CHello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                          Hello                                |
   *  //    |                                                               |
   *
   *
   *  worksheet.set_header("&RHello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                                                         Hello |
   *  //    |                                                               |
   * @endcode
   *
   * For simple text, if you do not specify any justification the text will be
   * centered. However, you must prefix the text with `&C` if you specify a font
   * name or any other formatting:
   *
   * @code
   *  worksheet.set_header("Hello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                          Hello                                |
   *  //    |                                                               |
   * @endcode
   *
   * You can have text in each of the justification regions:
   *
   * @code
   *  worksheet.set_header("&LCiao&CBello&RCielo");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    | Ciao                     Bello                          Cielo |
   *  //    |                                                               |
   * @endcode
   *
   * The information control characters act as variables that Excel will update
   * as the workbook or worksheet changes. Times and dates are in the users
   * default format:
   *
   * @code
   *  worksheet.set_header("&CPage &P of &N");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                        Page 1 of 6                            |
   *  //    |                                                               |
   *
   *  worksheet.set_header("&CUpdated at &T");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                    Updated at 12:30 PM                        |
   *  //    |                                                               |
   * @endcode
   *
   * You can specify the font size of a section of the text by prefixing it with
   * the control character `&n` where `n` is the font size:
   *
   * @code
   *  worksheet1.set_header("&C&30Hello Big");
   *  worksheet2.set_header("&C&10Hello Small");
   * @endcode
   *
   * You can specify the font of a section of the text by prefixing it with the
   * control sequence `&"font,style"` where `fontname` is a font name such as
   * Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
   * "Courier New" or "Times New Roman" and `style` is one of the standard
   *
   * @code
   *  worksheet1.set_header("&C&\"Courier New,Italic\"Hello");
   *  worksheet2.set_header("&C&\"Courier New,Bold Italic\"Hello");
   *  worksheet3.set_header("&C&\"Times New Roman,Regular\"Hello");
   * @endcode
   *
   * It is possible to combine all of these features together to create
   * sophisticated headers and footers. As an aid to setting up complicated
   * headers and footers you can record a page set-up as a macro in Excel and
   * look at the format strings that VBA produces. Remember however that VBA
   * uses two double quotes `""` to indicate a single double quote. For the last
   * example above the equivalent VBA code looks like this:
   *
   * @code
   *  .LeftHeader = ""
   *  .CenterHeader = "&""Times New Roman,Regular""Hello"
   *  .RightHeader = ""
   * @endcode
   *
   * Alternatively you can inspect the header and footer strings in an Excel
   * file by unzipping it and grepping the XML sub-files. The following shows
   * how to do that using libxml's xmllint to format the XML for clarity:
   *
   * @code
   *  $ unzip myfile.xlsm -d myfile
   *  $ xmllint --format `find myfile -name "*.xml" | xargs` | egrep "Header|Footer" | sed 's/&amp;/\&/g'
   *
   *  <headerFooter scaleWithDoc="0">
   *    <oddHeader>&L&P</oddHeader>
   *  </headerFooter>
   * @endcode
   *
   * To include a single literal ampersand `&` in a header or footer you should
   * use a double ampersand `&&`:
   *
   * @code
   *  worksheet.set_header("&CCuriouser && Curiouser - Attorneys at Law");
   * @endcode
   *
   * @note Excel requires that the header or footer string cannot be longer than 255
   * characters, including the control characters. Strings longer than this will
   * not be written.
   *
   * The `header_footer_options_t` options are:
   *
   * - `%margin_`: Header or footer margin in inches. The value must by larger
   *   than 0.0. The Excel default is 0.3.
   *
   * - `%image_left_`: The left header image filename, with path if required. This
   *   should have a corresponding `&G` placeholder in the `&L`
   *   section of the header/footer string.
   *
   * - `%image_center_`: The center header image filename, with path if
   *   required. This should have a corresponding `&G` placeholder in
   *   the `&C` section of the header/footer string.
   *
   * - `%image_right_`: The right header image filename, with path if
   *   required. This should have a corresponding `&G` placeholder in
   *   the `&R` section of the header/footer string.
   *
   * @code
   *  xwpp::header_footer_options_t header_options{ .margin_ = 0.2 };
   *
   *  worksheet.set_header("Some text", header_options);
   * @endcode
   *
   * Images can be inserted in the header by specifying the `&G`
   * placeholder and a filename/path to the image:
   *
   * @code
   *  xwpp::header_footer_options_t header_options{.image_left_ = "logo.png"};
   *
   *  worksheet1set_header("&L&G", header_options);
   * @endcode
   *
   * @image html headers_footers.png
   */
  void set_header(const std::string& str, const std::optional<header_footer_options_t>& options);
  /// @overload
  void set_header(const std::string& str);

  /**
   * @brief Set the printed page footer caption with additional options.
   *
   * @param str     The footer string.
   * @param options Footer options.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The syntax of this function is the same as `set_header()`.
   */
  void set_footer(const std::string& str, const std::optional<header_footer_options_t>& options);
  /// @overload
  void set_footer(const std::string& str);

  /**
   * @brief Set the sheet margins for the printed page.
   *
   * @param left   Left margin in inches.
   * @param right  Right margin in inches.
   * @param top    Top margin in inches.
   * @param bottom Bottom margin in inches.
   *
   * The `%set_margins()` function is used to set the margins of the
   * sheet when it is printed. The units are in inches. Specifying `-1` for
   * any parameter will give the default Excel value as shown above.
   *
   * @code
   *  worksheet.set_margins(1.3, 1.2, -1, -1);
   * @endcode
   *
   * @todo Use a structure?
   *
   * @note The Excel defaults are:
   *  - Left and right margins: 0.7.
   *  - Top and bottem margins: 0.75.
   */
  void set_margins(double left, double right, double top, double bottom);

  /**
   * @brief Set the page orientation as landscape.
   *
   * This function is used to set the orientation of a sheet's
   * printed page to landscape:
   *
   * @code
   *  worksheet.set_landscape();
   * @endcode
   *
   * @note The default orientation are:
   *  - Portrait for worksheet.
   *  - Landscape for chartsheet.
   */
  void set_landscape();

  /**
   * @brief Set the page orientation as portrait.
   *
   * This function is used to set the orientation of a sheet's
   * printed page to portrait:
   *
   * @code
   *  worksheet.set_portrait();
   * @endcode
   *
   * @note The default orientation are:
   *  - Portrait for worksheet.
   *  - Landscape for chartsheet.
   */
  void set_portrait();

  /**
   * @brief Hide the current sheet.
   *
   * The `%hide()` function is used to hide a sheet:
   *
   * @code
   *  worksheet.hide();
   * @endcode
   *
   * You may wish to hide a sheet in order to avoid confusing a user with
   * intermediate data or calculations.
   *
   * @image html hide_sheet.png
   *
   * A hidden sheet can not be activated or selected so this function is
   * mutually exclusive with the `activate()` and `select()` functions.
   * In addition, since the first sheet will default to being the
   * active worksheet, you cannot hide the first worksheet without activating
   * another sheet:
   *
   * @code
   *  worksheet2.activate(worksheet2);
   *  worksheet1.hide(worksheet1);
   * @endcode
   */
  void hide();

  /**
   * @brief Make a sheet the active, i.e., visible sheet.
   *
   * The `%activate()` function is used to specify which sheet is
   * initially visible in a multi-sheet workbook:
   *
   * @code
   *  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
   *  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();
   *  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();
   *
   *  worksheet3.activate();
   * @endcode
   *
   * @image html worksheet_activate.png
   *
   * More than one sheet can be selected via the `select()`
   * function, see below, however only one worksheet can be active.
   *
   * The default active worksheet is the first sheet.
   */
  void activate();

  /**
   * @brief Set a sheet tab as selected.
   *
   * The `%select()` function is used to indicate that a sheet is
   * selected in a multi-sheet workbook:
   *
   * @code
   *  worksheet1.activate();
   *  worksheet2.select();
   *  worksheet3.select();
   * @endcode
   *
   * A selected sheet has its tab highlighted. Selecting sheets is a
   * way of grouping them together so that, for example, several sheets
   * could be printed in one go. A sheet that has been activated via the
   * `activate()` function will also appear as selected.
   */
  void select();

  /**
   * @brief Set current sheet as the first visible sheet tab.
   *
   * The `activate()` function determines which sheet is initially
   * selected.  However, if there are a large number of sheets the selected
   * sheet may not appear on the screen. To avoid this you can select the
   * leftmost visible sheet tab using `%set_first_sheet()`:
   *
   * @code
   *  worksheet19.set_first_sheet(); // First visible worksheet tab.
   *  worksheet20.activate();        // First visible worksheet.
   * @endcode
   *
   * This function is not required very often. The default value is the first
   * sheet.
   */
  void set_first_sheet();

  void set_dpi(uint16_t horizontal_dpi, uint16_t vertical_dpi);

  // Retrieve function used by workbook and packager.
  [[nodiscard]] const std::string& get_sheet_name() const;
  [[nodiscard]] size_t get_sheet_index() const;
  [[nodiscard]] bool is_chartsheet() const;
  [[nodiscard]] bool is_hidden() const;
  [[nodiscard]] virtual bool has_background_image() const;
  [[nodiscard]] virtual object_properties_t get_background_image() const;
  [[nodiscard]] virtual bool has_embedded_image() const;
  [[nodiscard]] virtual std::vector<object_properties_t>& get_embedded_image_properties();
  [[nodiscard]] virtual const std::vector<object_properties_t>& get_embedded_image_properties() const;

  // Copy/move constructor and assignment operator are `deleted` so
  // reference return by `add_worksheet()`/`add_chartsheet()` cannot by
  // dereferenced by error.
  sheet_t(sheet_t&)             = delete;
  sheet_t& operator=(sheet_t)   = delete;
  sheet_t(sheet_t&&)            = delete;
  sheet_t& operator=(sheet_t&&) = delete;

  static const row_num_t ROW_MAX;
  static const col_num_t COL_MAX;

protected:
  sheet_t(bool is_chartsheet, const sheet_init_data_t& init_data);

  // Helper functions.
  void push_chart(const object_properties_t& object_props);
  static void get_image_properties(object_properties_t& image_props);
  void set_header_footer_image(const std::string& filename, image_position_t image_position);
  void add_drawing_object(const drawing_object_t& drawing_object, uint32_t drawing_id, bool chartsheet);
  void add_chart_link(uint32_t chart_ref_id);
  [[nodiscard]] uint32_t add_link_get_index(const std::string& url,
                                            const std::tuple<std::string, std::string, std::string>& relation);
  [[nodiscard]] uint32_t add_image_link_get_index(uint32_t image_ref_id, const object_properties_t& object_props);
  [[nodiscard]] uint32_t get_drawing_rel_index(const std::string& target);

  // Getter/setter for properties
  [[nodiscard]] virtual bool is_outline_changed() const;
  [[nodiscard]] bool is_selected() const;
  [[nodiscard]] uint16_t get_zoom() const;
  [[nodiscard]] color_t get_tab_color() const;
  [[nodiscard]] uint8_t get_paper_size() const;
  [[nodiscard]] bool is_page_setup_changed() const;
  void set_page_setup_changed(bool page_setup_changed);
  [[nodiscard]] drawing_orientation_t get_orientation() const;
  [[nodiscard]] uint16_t get_horizontal_dpi() const;
  [[nodiscard]] uint16_t get_vertical_dpi() const;
  void set_protection_obj(const protection_obj_t& protection);
  void add_image_properties(const object_properties_t& image_properties);
  [[nodiscard]] bool has_header_vml() const;
  [[nodiscard]] virtual size_t get_table_count() const;
  [[nodiscard]] uint16_t rel_count_next_value();

  // Functions to generate XML.
  [[nodiscard]] std::string write_tab_color() const;
  [[nodiscard]] virtual std::string write_sheet_pr() const = 0;
  [[nodiscard]] virtual std::string write_sheet_view()     = 0;
  [[nodiscard]] std::string write_sheet_views();
  [[nodiscard]] std::string write_sheet_protection() const;
  [[nodiscard]] std::string write_page_margins() const;
  [[nodiscard]] virtual std::string write_page_setup() const = 0;
  [[nodiscard]] std::string write_odd_header() const;
  [[nodiscard]] std::string write_odd_footer() const;
  [[nodiscard]] std::string write_header_footer() const;
  [[nodiscard]] std::string write_drawings();

private:
  // TODO Rework
  friend class packager_t;
  friend class workbook_t;

  // Preparation functions called by workbook and/or packager.
  void virtual set_error_cell(const object_properties_t& object_props, uint32_t ref_id)     = 0;
  void virtual prepare_background(uint32_t image_ref_id, object_properties_t& object_props) = 0;
  [[nodiscard]] uint32_t find_drawing_rel_index(const std::string& target);
  virtual void prepare_image(uint32_t image_ref_id, uint32_t drawing_id, object_properties_t& object_props) = 0;
  virtual void prepare_chart(uint32_t chart_ref_id, uint32_t drawing_id, object_properties_t& object_props) = 0;
  [[nodiscard]] uint32_t find_vml_drawing_rel_index(const std::string& target);
  [[nodiscard]] uint32_t get_vml_drawing_rel_index(const std::string& target);
  void prepare_header_image(uint32_t image_ref_id, object_properties_t& object_props);

  // Properties of sheet
  std::string name_;
  std::string quoted_name_;
  color_t tab_color_{color_t::UNSET};
  size_t index_{0};
  bool is_chartsheet_{false};
  bool hidden_{false};
  bool selected_{false};

  // Content
  uint16_t rel_count_{0};
  std::vector<object_properties_t> chart_data_;

  // Zoom properties
  uint16_t zoom_{100};

  // Print options
  uint16_t horizontal_dpi_{0};
  uint16_t vertical_dpi_{0};
  drawing_orientation_t orientation_{drawing_orientation_t::PORTRAIT};

  // Header footer
  bool header_footer_changed_{false};
  bool has_header_vml_{false};
  std::string header_;
  std::string footer_;
  std::vector<vml_obj_t> header_image_objs_;
  std::array<std::optional<object_properties_t>, 6> header_footer_objs_{std::nullopt, std::nullopt, std::nullopt,
                                                                        std::nullopt, std::nullopt, std::nullopt};

  // Page setup
  bool page_setup_changed_{false};
  uint8_t paper_size_{0};
  double margin_left_{0.7};
  double margin_right_{0.7};
  double margin_top_{0.75};
  double margin_bottom_{0.75};
  double margin_header_{0.3};
  double margin_footer_{0.3};
  protection_obj_t protection_;

  // Image
  std::vector<object_properties_t> image_props_;
  std::optional<drawing_t> drawing_;
  std::map<std::string, uint32_t, std::less<>> vml_drawing_rel_ids_;
  uint32_t vml_drawing_rel_id_{0};
  std::map<std::string, uint32_t, std::less<>> drawing_rel_ids_;
  uint32_t drawing_rel_id_{0};

  // Links
  std::vector<std::tuple<std::string, std::string, std::string>> external_drawing_links_;
  std::vector<std::tuple<std::string, std::string, std::string>> drawing_links_;
  std::vector<std::tuple<std::string, std::string, std::string>> vml_drawing_links_;

  // Access to workbook sheets properties
  size_t* active_sheet_{nullptr};
  size_t* first_sheet_{nullptr};
};

/// @cond
}
/// @endcond

#endif
