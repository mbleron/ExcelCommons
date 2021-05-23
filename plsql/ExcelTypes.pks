create or replace package ExcelTypes is

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
    name     varchar2(64)
  , b        boolean := false
  , i        boolean := false
  , color    varchar2(8)
  , sz       pls_integer
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
  , content     varchar2(32767)
  );
  
  function isValidPatternType (p_patternType in varchar2) return boolean;
  function isValidBorderStyle (p_borderStyle in varchar2) return boolean;
  function isValidHorizontalAlignment (p_hAlignment in varchar2) return boolean;
  function isValidVerticalAlignment (p_vAlignment in varchar2) return boolean;
  
  function getFillPatternTypeId (p_patternType in varchar2) return pls_integer;
  function getBorderStyleId (p_borderStyle in varchar2) return pls_integer;
  function getHorizontalAlignmentId (p_hAlignment in varchar2) return pls_integer;
  function getVerticalAlignmentId (p_vAlignment in varchar2) return pls_integer;

end;
/
