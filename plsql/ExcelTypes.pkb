create or replace package body ExcelTypes is

  type simpleTypeMap_t is table of pls_integer index by varchar2(16);
  
  underlineStyleMap  simpleTypeMap_t;
  fillPatternMap     simpleTypeMap_t;
  borderStyleMap     simpleTypeMap_t;
  hAlignmentMap      simpleTypeMap_t;
  vAlignmentMap      simpleTypeMap_t;
  
  procedure initialize
  is
  begin
    
    -- underline style
    underlineStyleMap('none') := 0;
    underlineStyleMap('single') := 1;
    underlineStyleMap('double') := 2;
    underlineStyleMap('singleAccounting') := 33;
    underlineStyleMap('doubleAccounting') := 34;
    
    -- fill pattern type
    fillPatternMap('none') := 0;
    fillPatternMap('solid') := 1;
    fillPatternMap('mediumGray') := 2;
    fillPatternMap('darkGray') := 3;
    fillPatternMap('lightGray') := 4;
    fillPatternMap('darkHorizontal') := 5;
    fillPatternMap('darkVertical') := 6;
    fillPatternMap('darkDown') := 7;
    fillPatternMap('darkUp') := 8;
    fillPatternMap('darkGrid') := 9;
    fillPatternMap('darkTrellis') := 10;
    fillPatternMap('lightHorizontal') := 11;
    fillPatternMap('lightVertical') := 12;
    fillPatternMap('lightDown') := 13;
    fillPatternMap('lightUp') := 14;
    fillPatternMap('lightGrid') := 15;
    fillPatternMap('lightTrellis') := 16;
    fillPatternMap('gray125') := 17;
    fillPatternMap('gray0625') := 18;
    
    -- border style
    borderStyleMap('none') := 0;
    borderStyleMap('thin') := 1;
    borderStyleMap('medium') := 2;
    borderStyleMap('dashed') := 3;
    borderStyleMap('dotted') := 4;
    borderStyleMap('thick') := 5;
    borderStyleMap('double') := 6;
    borderStyleMap('hair') := 7;
    borderStyleMap('mediumDashed') := 8;
    borderStyleMap('dashDot') := 9;
    borderStyleMap('mediumDashDot') := 10;
    borderStyleMap('dashDotDot') := 11;
    borderStyleMap('mediumDashDotDot') := 12;
    borderStyleMap('slantDashDot') := 13;
    
    -- horizontal alignment
    hAlignmentMap('general') := 0;
    hAlignmentMap('left') := 1;
    hAlignmentMap('center') := 2;
    hAlignmentMap('right') := 3;
    hAlignmentMap('fill') := 4;
    hAlignmentMap('justify') := 5;
    hAlignmentMap('centerContinuous') := 6;
    hAlignmentMap('distributed') := 7;
    
    -- vertical alignment
    vAlignmentMap('top') := 0;
    vAlignmentMap('center') := 1;
    vAlignmentMap('bottom') := 2;
    vAlignmentMap('justify') := 3;
    vAlignmentMap('distributed') := 4;
    
  end;
  
  function isValidUnderlineStyle (p_underlineStyle in varchar2) return boolean is
  begin
    return underlineStyleMap.exists(p_underlineStyle);
  end;
  
  function isValidPatternType (p_patternType in varchar2) return boolean is
  begin
    return fillPatternMap.exists(p_patternType);
  end;

  function isValidBorderStyle (p_borderStyle in varchar2) return boolean is
  begin
    return borderStyleMap.exists(p_borderStyle);
  end;

  function isValidHorizontalAlignment (p_hAlignment in varchar2) return boolean is
  begin
    return hAlignmentMap.exists(p_hAlignment);
  end;

  function isValidVerticalAlignment (p_vAlignment in varchar2) return boolean is
  begin
    return vAlignmentMap.exists(p_vAlignment);
  end;

  function getUnderlineStyleId (p_underlineStyle in varchar2) return pls_integer is
  begin  
    return underlineStyleMap(p_underlineStyle);
  end;
  
  function getFillPatternTypeId (p_patternType in varchar2) return pls_integer is
  begin
    return fillPatternMap(p_patternType);
  end;

  function getBorderStyleId (p_borderStyle in varchar2) return pls_integer is
  begin
    return borderStyleMap(p_borderStyle);
  end;
  
  function getHorizontalAlignmentId (p_hAlignment in varchar2) return pls_integer is
  begin
    return hAlignmentMap(p_hAlignment);
  end;

  function getVerticalAlignmentId (p_vAlignment in varchar2) return pls_integer is
  begin
    return vAlignmentMap(p_vAlignment);
  end;

begin
  
  initialize;

end;
/
