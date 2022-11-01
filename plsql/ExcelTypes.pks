create or replace package ExcelTypes is
/* ======================================================================================

  MIT License

  Copyright (c) 2021-2022 Marc Bleron

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.

=========================================================================================
    Change history :
    Marc Bleron       2021-05-23     Creation
    Marc Bleron       2021-09-04     Added wrapText attribute, and font underline
    Marc Bleron       2022-09-03     Added CSS features
====================================================================================== */

  DEFAULT_FONT_FAMILY   constant varchar2(256) := 'Calibri';
  DEFAULT_FONT_SIZE     constant number := 11; -- points

  subtype uint8 is pls_integer range 0..255;

  type CT_BorderPr is record (
    style  varchar2(16)
  , color  varchar2(8)
  );
  
  type CT_Border is record (
    left     CT_BorderPr
  , right    CT_BorderPr
  , top      CT_BorderPr
  , bottom   CT_BorderPr
  , content  varchar2(32767)
  );
  
  type CT_Font is record (
    name     varchar2(64) := DEFAULT_FONT_FAMILY
  , b        boolean := false
  , i        boolean := false
  , u        varchar2(16)
  , color    varchar2(8)
  , sz       pls_integer := DEFAULT_FONT_SIZE
  , content  varchar2(32767)
  );

  type CT_PatternFill is record (
    patternType  varchar2(32)
  , fgColor      varchar2(8)
  , bgColor      varchar2(8)
  );

  type CT_Fill is record (
    patternFill  CT_PatternFill
  , content      varchar2(32767)
  );
  
  type CT_CellAlignment is record (
    horizontal  varchar2(16)
  , vertical    varchar2(16)
  , wrapText    boolean
  , content     varchar2(32767)
  );
  
  type CT_Style is record (
    numberFormat  varchar2(256)
  , font          CT_Font
  , fill          CT_Fill
  , border        CT_Border
  , alignment     CT_CellAlignment
  );
  
  type colorMap_t is table of varchar2(6) index by varchar2(20);
  function getColorMap return colorMap_t;
  
  function isValidColorName (p_colorName in varchar2) return boolean;
  function isValidUnderlineStyle (p_underlineStyle in varchar2) return boolean;
  function isValidPatternType (p_patternType in varchar2) return boolean;
  function isValidBorderStyle (p_borderStyle in varchar2) return boolean;
  function isValidHorizontalAlignment (p_hAlignment in varchar2) return boolean;
  function isValidVerticalAlignment (p_vAlignment in varchar2) return boolean;
  
  function makeRgbColor (r in uint8, g in uint8, b in uint8, a in number default null) return varchar2;
  function validateColor (colorSpec in varchar2) return varchar2;
  function getUnderlineStyleId (p_underlineStyle in varchar2) return pls_integer;
  function getFillPatternTypeId (p_patternType in varchar2) return pls_integer;
  function getBorderStyleId (p_borderStyle in varchar2) return pls_integer;
  function getHorizontalAlignmentId (p_hAlignment in varchar2) return pls_integer;
  function getVerticalAlignmentId (p_vAlignment in varchar2) return pls_integer;

  function makeBorderPr (
    p_style  in varchar2 default null
  , p_color  in varchar2 default null
  )
  return CT_BorderPr;
  
  function makeBorder (
    p_left    in CT_BorderPr default makeBorderPr()
  , p_right   in CT_BorderPr default makeBorderPr()
  , p_top     in CT_BorderPr default makeBorderPr()
  , p_bottom  in CT_BorderPr default makeBorderPr()
  )
  return CT_Border;

  function makeBorder (
    p_style  in varchar2
  , p_color  in varchar2 default null
  )
  return CT_Border;

  function makeFont (
    p_name   in varchar2 default null
  , p_sz     in pls_integer default null
  , p_b      in boolean default false
  , p_i      in boolean default false
  , p_color  in varchar2 default null
  , p_u      in varchar2 default null
  )
  return CT_Font;
  
  function makePatternFill (
    p_patternType  in varchar2
  , p_fgColor      in varchar2 default null
  , p_bgColor      in varchar2 default null
  )
  return CT_Fill;

  function makeAlignment (
    p_horizontal  in varchar2 default null
  , p_vertical    in varchar2 default null
  , p_wrapText    in boolean default false
  )
  return CT_CellAlignment;
  
  function mergeBorders (masterBorder in CT_Border, border in CT_Border) return CT_Border;
  function mergeFonts (masterFont in CT_Font, font in CT_Font) return CT_Font;
  function mergePatternFills (masterFill in CT_Fill, fill in CT_Fill) return CT_Fill;
  function mergeAlignments (masterAlignment in CT_CellAlignment, alignment in CT_CellAlignment) return CT_CellAlignment;
  
  function getStyleFromCss (cssString in varchar2) return CT_Style;
  
  procedure testCss (cssString in varchar2);

end;
/
