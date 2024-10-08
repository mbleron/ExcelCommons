create or replace package ExcelTypes is
/* ======================================================================================

  MIT License

  Copyright (c) 2021-2024 Marc Bleron

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
    Marc Bleron       2022-11-04     Added gradientFill
    Marc Bleron       2023-02-15     Added font vertical alignment (super/sub-script)
                                     and rich text support
    Marc Bleron       2024-02-23     Fix: NLS-independent conversion of CSS number-token
                                     Added font strikethrough, text orientation, indent
    Marc Bleron       2024-03-13     Added definedName structure
    Marc Bleron       2024-08-13     Added dataValidation structure
====================================================================================== */

  DEFAULT_FONT_FAMILY   constant varchar2(256) := 'Calibri';
  DEFAULT_FONT_SIZE     constant number := 11; -- points
  FT_PATTERN            constant pls_integer := 0;
  FT_GRADIENT           constant pls_integer := 1;

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
    name       varchar2(64)
  , b          boolean := false
  , i          boolean := false
  , u          varchar2(16)
  , color      varchar2(8)
  , sz         pls_integer
  , vertAlign  varchar2(16)
  , strike     boolean := false
  , content    varchar2(32767)
  );

  type CT_PatternFill is record (
    patternType  varchar2(32)
  , fgColor      varchar2(8)
  , bgColor      varchar2(8)
  );
  
  type CT_GradientStop is record (
    position  number
  , color     varchar2(8)
  );

  type CT_GradientStopList is table of CT_GradientStop;
  
  type CT_GradientFill is record (
    degree  number
  , stops   CT_GradientStopList
  );
  
  type CT_NumFmt is record (
    numFmtId     pls_integer
  , formatCode   varchar2(256)
  , isDate       boolean := false
  , isTimestamp  boolean := false
  );
  
  type CT_NumFmtMap is table of CT_NumFmt index by pls_integer;

  type CT_Fill is record (
    patternFill   CT_PatternFill
  , gradientFill  CT_GradientFill
  , content       varchar2(32767)
  , fillType      pls_integer
  );
  
  type CT_CellAlignment is record (
    horizontal    varchar2(16)
  , vertical      varchar2(16)
  , wrapText      boolean
  , textRotation  number
  , verticalText  boolean := false
  , indent        number
  , content       varchar2(32767)
  );
  
  type CT_Style is record (
    numberFormat  varchar2(256)
  , font          CT_Font
  , fill          CT_Fill
  , border        CT_Border
  , alignment     CT_CellAlignment
  );
  
  type CT_TextRun is record (
    text  varchar2(32767)
  , font  CT_Font
  );
  
  type CT_TextRunList is table of CT_TextRun;
  
  type CT_RichText is record (
    runs     CT_TextRunList
  , content  varchar2(32767)
  );

  type CT_SheetBase is record (
    idx   pls_integer  
  , name  varchar2(31 char)
  );
  
  type CT_Sheets is table of CT_SheetBase;
  
  type CT_DefinedName is record (
    idx             pls_integer
  , name            varchar2(255 char)
  , scope           varchar2(128)
  , formula         varchar2(32767)
  , cellRef         varchar2(10)
  , refStyle        pls_integer
  , comment         varchar2(255 char)
  , hidden          boolean
  , futureFunction  boolean
  , builtIn         boolean
  );
  
  type CT_DefinedNames is table of CT_DefinedName;
  type CT_DefinedNameMap is table of CT_DefinedName index by varchar2(2048);
  
  type supportingLinks_t is table of pls_integer index by pls_integer;
  
  type xti_t is record (
    idx           pls_integer
  , externalLink  pls_integer
  , firstSheet    pls_integer
  , lastSheet     pls_integer
  );
  
  type xtiMap_t is table of xti_t index by varchar2(24);
  type xtiArray_t is table of xti_t;
  
  type CT_Externals is record (
    supLinks  supportingLinks_t
  , xtiArray  xtiArray_t
  , xtiMap    xtiMap_t
  );
  
  subtype ST_Ref is varchar2(128);
  type ST_Sqref is table of ST_Ref;
  
  type cellRange_t is record (
    value     ST_Ref
  , rwFirst   pls_integer
  , rwLast    pls_integer
  , colFirst  pls_integer
  , colLast   pls_integer
  );
  
  type cellRangeList_t is table of cellRange_t;
  
  type CT_DataValidation is record (
    allowBlank        boolean
  , error             varchar2(225 char)
  , errorStyle        varchar2(128)
  , errorTitle        varchar2(32 char)
  , operator          varchar2(128)
  , prompt            varchar2(255 char)
  , promptTitle       varchar2(32 char)
  , showDropDown      boolean
  , showErrorMessage  boolean
  , showInputMessage  boolean
  , sqref             cellRangeList_t
  , activeCellRef     varchar2(10)
  , type              varchar2(128)
  , fmla1             varchar2(8192)
  , fmla2             varchar2(8192)
  , refStyle1         pls_integer
  , refStyle2         pls_integer
  , internalCellRef   varchar2(10)
  );
  
  type CT_DataValidations is table of CT_DataValidation;
  
  type colorMap_t is table of varchar2(6) index by varchar2(20);
  function getColorMap return colorMap_t;
  
  function isValidColorName (p_colorName in varchar2) return boolean;
  function isValidUnderlineStyle (p_underlineStyle in varchar2) return boolean;
  function isValidPatternType (p_patternType in varchar2) return boolean;
  function isValidBorderStyle (p_borderStyle in varchar2) return boolean;
  function isValidHorizontalAlignment (p_hAlignment in varchar2) return boolean;
  function isValidVerticalAlignment (p_vAlignment in varchar2) return boolean;
  function isValidFontVerticalAlignment (p_fontVertAlign in varchar2) return boolean;
  function isValidDataValidationType (p_dataValType in varchar2) return boolean;
  function isValidDataValidationOperator (p_dataValOp in varchar2) return boolean;
  function isValidDataValidationErrStyle (p_dataValErrStyle in varchar2) return boolean;
  
  function makeRgbColor (r in uint8, g in uint8, b in uint8, a in number default null) return varchar2;
  function validateColor (colorSpec in varchar2) return varchar2;
  function getUnderlineStyleId (p_underlineStyle in varchar2) return pls_integer;
  function getFillPatternTypeId (p_patternType in varchar2) return pls_integer;
  function getBorderStyleId (p_borderStyle in varchar2) return pls_integer;
  function getHorizontalAlignmentId (p_hAlignment in varchar2) return pls_integer;
  function getVerticalAlignmentId (p_vAlignment in varchar2) return pls_integer;
  function getFontVerticalAlignmentId (p_fontVertAlign in varchar2) return pls_integer;
  function getDataValidationTypeId (p_dataValType in varchar2) return pls_integer;
  function getDataValidationOpId (p_dataValOp in varchar2) return pls_integer;
  function getDataValidationErrStyleId (p_dataValErrStyle in varchar2) return pls_integer;

  function makeNumFmt (
    numFmtId in pls_integer
  , formatCode in varchar2
  ) 
  return CT_NumFmt;

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
    p_name       in varchar2 default null
  , p_sz         in pls_integer default null
  , p_b          in boolean default false
  , p_i          in boolean default false
  , p_color      in varchar2 default null
  , p_u          in varchar2 default null
  , p_vertAlign  in varchar2 default null
  , p_strike     in boolean default false
  )
  return CT_Font;
  
  function makePatternFill (
    p_patternType  in varchar2
  , p_fgColor      in varchar2 default null
  , p_bgColor      in varchar2 default null
  )
  return CT_Fill;

  function makeGradientStop (
    p_position  in number
  , p_color     in varchar2
  )
  return CT_GradientStop;
  
  function makeGradientFill (
    p_degree  in number default null
  , p_stops   in CT_GradientStopList default null
  )
  return CT_Fill;

  procedure addGradientStop (
    p_fill      in out nocopy CT_Fill
  , p_position  in number
  , p_color     in varchar2
  );

  function makeAlignment (
    p_horizontal    in varchar2 default null
  , p_vertical      in varchar2 default null
  , p_wrapText      in boolean default false
  , p_textRotation  in number default null
  , p_verticalText  in boolean default false
  , p_indent        in number default null
  )
  return CT_CellAlignment;

  function makeRichText (
    p_content   in xmltype
  , p_rootFont  in CT_Font
  )
  return CT_RichText;
  
  function mergeBorders (masterBorder in CT_Border, border in CT_Border) return CT_Border;
  function mergeFonts (masterFont in CT_Font, font in CT_Font, force in boolean default false) return CT_Font;
  function mergePatternFills (masterFill in CT_Fill, fill in CT_Fill) return CT_Fill;
  function mergeAlignments (masterAlignment in CT_CellAlignment, alignment in CT_CellAlignment) return CT_CellAlignment;
  
  function applyBorderSide (border in CT_Border, top in boolean, right in boolean, bottom in boolean, left in boolean) return CT_Border;
  
  function getStyleFromCss (cssString in varchar2) return CT_Style;
  
  procedure testCss (cssString in varchar2);
  
  function fromOADate (p_value in number, p_scale in pls_integer default 0) return timestamp_unconstrained;
  function getBuiltInDateFmts return CT_NumFmtMap;
  
  function isSheetQuotableStartChar (p_char in varchar2) return boolean;
  function isSheetQuotableChar (p_char in varchar2) return boolean;
  function isNameStartChar (p_char in varchar2) return boolean;
  function isNameChar (p_char in varchar2) return boolean;
  
  procedure setDebug (p_status in boolean);

end;
/
