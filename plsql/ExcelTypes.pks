create or replace package ExcelTypes is
/* ======================================================================================

  MIT License

  Copyright (c) 2021-2025 Marc Bleron

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
    Marc Bleron       2024-09-04     Added conditionalFormatting structures
====================================================================================== */

  DEFAULT_FONT_FAMILY   constant varchar2(256) := 'Calibri';
  DEFAULT_FONT_SIZE     constant number := 11; -- points
  FT_PATTERN            constant pls_integer := 0;
  FT_GRADIENT           constant pls_integer := 1;

  -- 2.5.18 CFType
  CF_TYPE_CELLIS        constant pls_integer := 1;
  CF_TYPE_EXPRIS        constant pls_integer := 2;
  CF_TYPE_GRADIENT      constant pls_integer := 3;
  CF_TYPE_DATABAR       constant pls_integer := 4;
  CF_TYPE_FILTER        constant pls_integer := 5;
  CF_TYPE_MULTISTATE    constant pls_integer := 6;
  
  --CF_TYPE_CELLIS             constant pls_integer := 1;
  CF_TYPE_EXPR               constant pls_integer := 7;
  CF_TYPE_COLORSCALE         constant pls_integer := 8;
  --CF_TYPE_DATABAR            constant pls_integer := 4;
  CF_TYPE_ICONSET            constant pls_integer := 9;
  CF_TYPE_TOP                constant pls_integer := 10;
  CF_TYPE_BOTTOM             constant pls_integer := 11;
  CF_TYPE_UNIQUES            constant pls_integer := 12;
  CF_TYPE_DUPLICATES         constant pls_integer := 13;
  CF_TYPE_TEXT               constant pls_integer := 14;
  CF_TYPE_BLANKS             constant pls_integer := 15;
  CF_TYPE_NOBLANKS           constant pls_integer := 16;
  CF_TYPE_ERRORS             constant pls_integer := 17;
  CF_TYPE_NOERRORS           constant pls_integer := 18;
  CF_TYPE_TIMEPERIOD         constant pls_integer := 19;
  CF_TYPE_ABOVEAVERAGE       constant pls_integer := 20;
  CF_TYPE_BELOWAVERAGE       constant pls_integer := 21;
  CF_TYPE_EQUALABOVEAVERAGE  constant pls_integer := 22;
  CF_TYPE_EQUALBELOWAVERAGE  constant pls_integer := 23;
  
  -- 2.5.16 CFTemp
  CF_TEMP_EXPR                 constant pls_integer := 0;
  CF_TEMP_FMLA                 constant pls_integer := 1;
  CF_TEMP_GRADIENT             constant pls_integer := 2;
  CF_TEMP_DATABAR              constant pls_integer := 3;
  CF_TEMP_MULTISTATE           constant pls_integer := 4;
  CF_TEMP_FILTER               constant pls_integer := 5;
  CF_TEMP_UNIQUEVALUES         constant pls_integer := 7;
  CF_TEMP_CONTAINSTEXT         constant pls_integer := 8;
  CF_TEMP_CONTAINSBLANKS       constant pls_integer := 9;
  CF_TEMP_CONTAINSNOBLANKS     constant pls_integer := 10;
  CF_TEMP_CONTAINSERRORS       constant pls_integer := 11;
  CF_TEMP_CONTAINSNOERRORS     constant pls_integer := 12;
  CF_TEMP_TIMEPERIODTODAY      constant pls_integer := 15;
  CF_TEMP_TIMEPERIODTOMORROW   constant pls_integer := 16;
  CF_TEMP_TIMEPERIODYESTERDAY  constant pls_integer := 17;
  CF_TEMP_TIMEPERIODLAST7DAYS  constant pls_integer := 18;
  CF_TEMP_TIMEPERIODLASTMONTH  constant pls_integer := 19;
  CF_TEMP_TIMEPERIODNEXTMONTH  constant pls_integer := 20;
  CF_TEMP_TIMEPERIODTHISWEEK   constant pls_integer := 21;
  CF_TEMP_TIMEPERIODNEXTWEEK   constant pls_integer := 22;
  CF_TEMP_TIMEPERIODLASTWEEK   constant pls_integer := 23;
  CF_TEMP_TIMEPERIODTHISMONTH  constant pls_integer := 24;
  CF_TEMP_ABOVEAVERAGE         constant pls_integer := 25;
  CF_TEMP_BELOWAVERAGE         constant pls_integer := 26;
  CF_TEMP_DUPLICATEVALUES      constant pls_integer := 27;
  CF_TEMP_EQUALABOVEAVERAGE    constant pls_integer := 29;
  CF_TEMP_EQUALBELOWAVERAGE    constant pls_integer := 30;
  
  -- 2.5.15 CFOper
  CF_OPER_BN  constant pls_integer := 1;
  CF_OPER_NB  constant pls_integer := 2;
  CF_OPER_EQ  constant pls_integer := 3;
  CF_OPER_NE  constant pls_integer := 4;
  CF_OPER_GT  constant pls_integer := 5;
  CF_OPER_LT  constant pls_integer := 6;
  CF_OPER_GE  constant pls_integer := 7;
  CF_OPER_LE  constant pls_integer := 8;
  
  -- 2.5.17 CFTextOper
  CF_TEXTOPER_CONTAINS     constant pls_integer := 0;
  CF_TEXTOPER_NOTCONTAINS  constant pls_integer := 1;
  CF_TEXTOPER_BEGINSWITH   constant pls_integer := 2;
  CF_TEXTOPER_ENDSWITH     constant pls_integer := 3;
  
  -- 2.5.12 CFDateOper
  CF_TIMEPERIOD_TODAY      constant pls_integer := 0;
  CF_TIMEPERIOD_YESTERDAY  constant pls_integer := 1;
  CF_TIMEPERIOD_LAST7DAYS  constant pls_integer := 2;
  CF_TIMEPERIOD_THISWEEK   constant pls_integer := 3;
  CF_TIMEPERIOD_LASTWEEK   constant pls_integer := 4;
  CF_TIMEPERIOD_LASTMONTH  constant pls_integer := 5;
  CF_TIMEPERIOD_TOMORROW   constant pls_integer := 6;
  CF_TIMEPERIOD_NEXTWEEK   constant pls_integer := 7;
  CF_TIMEPERIOD_NEXTMONTH  constant pls_integer := 8;
  CF_TIMEPERIOD_THISMONTH  constant pls_integer := 9;
  
  -- 2.5.86 KPISets
  CF_ICONSET_3ARROWS          constant pls_integer := 0;
  CF_ICONSET_3ARROWSGRAY      constant pls_integer := 1;
  CF_ICONSET_3FLAGS           constant pls_integer := 2;
  CF_ICONSET_3TRAFFICLIGHTS1  constant pls_integer := 3;
  CF_ICONSET_3TRAFFICLIGHTS2  constant pls_integer := 4;
  CF_ICONSET_3SIGNS           constant pls_integer := 5;
  CF_ICONSET_3SYMBOLS         constant pls_integer := 6;
  CF_ICONSET_3SYMBOLS2        constant pls_integer := 7;
  CF_ICONSET_4ARROWS          constant pls_integer := 8;
  CF_ICONSET_4ARROWSGRAY      constant pls_integer := 9;
  CF_ICONSET_4REDTOBLACK      constant pls_integer := 10;
  CF_ICONSET_4RATING          constant pls_integer := 11;
  CF_ICONSET_4TRAFFICLIGHTS   constant pls_integer := 12;
  CF_ICONSET_5ARROWS          constant pls_integer := 13;
  CF_ICONSET_5ARROWSGRAY      constant pls_integer := 14;
  CF_ICONSET_5RATING          constant pls_integer := 15;
  CF_ICONSET_5QUARTERS        constant pls_integer := 16;

  -- 2.5.19 CFVOtype
  CFVO_NUM         constant pls_integer := 1;
  CFVO_MIN         constant pls_integer := 2;
  CFVO_MAX         constant pls_integer := 3;
  CFVO_PERCENT     constant pls_integer := 4;
  CFVO_PERCENTILE  constant pls_integer := 5;
  CFVO_FMLA        constant pls_integer := 7;

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
  , b          boolean
  , i          boolean
  , u          varchar2(16)
  , color      varchar2(8)
  , sz         pls_integer
  , vertAlign  varchar2(16)
  , strike     boolean
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
  , numFmtId      pls_integer
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
  
  type cellRangeSeq_t is record (
    ranges                    cellRangeList_t
  , lastRangeCellRef          varchar2(10)  -- top-left cell of the last range in the sequence
  , boundingAreaFirstCellRef  varchar2(10)  -- top-left cell of the bounding area of all ranges in the sequence
  );
  
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
  , sqref             cellRangeSeq_t
  , type              varchar2(128)
  , fmla1             varchar2(8192)
  , fmla2             varchar2(8192)
  , refStyle1         pls_integer
  , refStyle2         pls_integer
  );
  
  type CT_DataValidations is table of CT_DataValidation;
  
  type CT_Cfvo is record (
    type      pls_integer
  , value     varchar2(8192)
  , gte       boolean
  , color     varchar2(256)
  , refStyle  pls_integer
  );
  
  type CT_CfvoList is table of CT_Cfvo;
  
  type CT_CfRule is record (
    sqref      cellRangeSeq_t
  , type       pls_integer
  , template   pls_integer
  , dxfId      pls_integer
  , priority   pls_integer
  , param      pls_integer
  , stopTrue   boolean
  , bottom     boolean
  , percent    boolean
  , strParam   varchar2(255 char)
  , fmla1      varchar2(8192)
  , fmla2      varchar2(8192)
  , fmla3      varchar2(8192)
  , cfvoList   CT_CfvoList
  , hideValue  boolean
  , iconSet    pls_integer
  , reverse    boolean
  , refStyle1  pls_integer
  , refStyle2  pls_integer
  , refStyle3  pls_integer
  );
  
  type CT_CfRules is table of CT_CfRule;
  
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
  function isValidCondFmtRuleType (p_type in pls_integer) return boolean;
  function isValidCondFmtOperator (p_type in pls_integer, p_operator in pls_integer) return boolean;
  function isValidCondFmtVOType (p_type in pls_integer) return boolean;
  function isValidCondFmtIconSet (p_iconSet in pls_integer) return boolean;
  
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

  function getCondFmtRuleType (p_type in pls_integer, p_temp in pls_integer) return varchar2;
  function getCondFmtTimePeriod (p_timePeriod in pls_integer) return varchar2;
  function getCondFmtOperator (p_temp in pls_integer, p_operator in pls_integer) return varchar2;
  function getCondFmtIconSet (p_iconSet in pls_integer) return varchar2;
  function getCondFmtVOType (p_cfvoType in pls_integer) return varchar2;

  function makeCfvo (
    p_type      in pls_integer default null
  , p_value     in varchar2 default null
  , p_gte       in boolean default null
  , p_color     in varchar2 default null
  , p_refStyle  in pls_integer default null
  )
  return CT_Cfvo;

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
  , p_b          in boolean default null
  , p_i          in boolean default null
  , p_color      in varchar2 default null
  , p_u          in varchar2 default null
  , p_vertAlign  in varchar2 default null
  , p_strike     in boolean default null
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
  
  procedure swapPatternFillColors (fill in out nocopy CT_Fill);
  function mergeBorders (masterBorder in CT_Border, border in CT_Border) return CT_Border;
  function mergeFonts (masterFont in CT_Font, font in CT_Font/*, force in boolean default false*/) return CT_Font;
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
