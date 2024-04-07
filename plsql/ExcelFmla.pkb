create or replace package body ExcelFmla is

  PTG_ADD       constant pls_integer := 3;   -- 0x03 PtgAdd - Addition (+)
  PTG_SUB       constant pls_integer := 4;   -- 0x04 PtgSub - Subtraction (-)
  PTG_MUL       constant pls_integer := 5;   -- 0x05 PtgMul - Multiplication (*)
  PTG_DIV       constant pls_integer := 6;   -- 0x06 PtgDiv - Division (/)  
  PTG_POWER     constant pls_integer := 7;   -- 0x07 PtgPower - Exponentiation (^)
  PTG_CONCAT    constant pls_integer := 8;   -- 0x08 PtgConcat - Concatenation (&)
  PTG_LT        constant pls_integer := 9;   -- 0x09 PtgLt - Less-than (<)
  PTG_LE        constant pls_integer := 10;  -- 0x0A PtgLe - Less-than-or-equal-to (<=)
  PTG_EQ        constant pls_integer := 11;  -- 0x0B PtgEq - Equal-to (=)
  PTG_GE        constant pls_integer := 12;  -- 0x0C PtgGe - Greater-than-or-equal-to (>=)
  PTG_GT        constant pls_integer := 13;  -- 0x0D PtgGt - Greater-than (>)
  PTG_NE        constant pls_integer := 14;  -- 0x0E PtgNe - Not-equal-to (<>)
  PTG_ISECT     constant pls_integer := 15;  -- 0x0F PtgIsect - Binary intersection operator (<space>)
  PTG_UNION     constant pls_integer := 16;  -- 0x10 PtgUnion - Binary union operator (,)
  PTG_RANGE     constant pls_integer := 17;  -- 0x11 PtgRange - Binary range operator (:)
  PTG_UPLUS     constant pls_integer := 18;  -- 0x12 PtgUPlus - Unary plus (+)
  PTG_UMINUS    constant pls_integer := 19;  -- 0x13 PtgUMinus - Unary minus (-)
  PTG_PERCENT   constant pls_integer := 20;  -- 0x14 PtgPercent - Percentage (%)    
  PTG_PAREN     constant pls_integer := 21;  -- 0x15 PtgParen - Parentheses
  PTG_MISSARG   constant pls_integer := 22;  -- 0x16 PtgMissArg - Missing value
  PTG_STR       constant pls_integer := 23;  -- 0x17 PtgString - String
  
  PTG_ATTR_B    constant pls_integer := 25;  -- 0x19 PtgAttr* 1st byte
  PTG_ATTRSEMI  constant pls_integer := 01;  -- 0x01 PtgAttrSemi 2nd byte
  PTG_ATTRSUM   constant pls_integer := 16;  -- 0x10 PtgAttrSum 2nd byte
  
  PTG_ERR       constant pls_integer := 28;  -- 0x1C PtgErr - Error value
  PTG_BOOL      constant pls_integer := 29;  -- 0x1D PtgBool - Boolean value 0/1
  PTG_INT       constant pls_integer := 30;  -- 0x1E PtgInt - 16bit unsigned integer
  PTG_NUM       constant pls_integer := 31;  -- 0x1F PtgNum - 64bit floating-point number
  
  -- Reference-type Ptg's
  PTG_ARRAY_R     constant pls_integer := 32;  -- 0x20 PtgArray
  PTG_FUNC_R      constant pls_integer := 33;  -- 0x21 PtgFunc
  PTG_FUNCVAR_R   constant pls_integer := 34;  -- 0x22 PtgFuncVar
  PTG_NAME_R      constant pls_integer := 35;  -- 0x23 PtgName
  PTG_REF_R       constant pls_integer := 36;  -- 0x24 PtgRef
  PTG_AREA_R      constant pls_integer := 37;  -- 0x25 PtgArea
  PTG_MEMAREA_R   constant pls_integer := 38;  -- 0x26 PtgMemArea
  PTG_MEMERR_R    constant pls_integer := 39;  -- 0x27 PtgMemErr
  PTG_MEMFUNC_R   constant pls_integer := 41;  -- 0x29 PtgMemFunc
  PTG_NAMEX_R     constant pls_integer := 57;  -- 0x39 PtgNameX
  PTG_REF3D_R     constant pls_integer := 58;  -- 0x3A PtgRef3d
  PTG_AREA3D_R    constant pls_integer := 59;  -- 0x3B PtgArea3d
  
  -- Value-type Ptg's
  PTG_ARRAY_V     constant pls_integer := 64;  -- 0x40 PtgArray
  PTG_FUNC_V      constant pls_integer := 65;  -- 0x41 PtgFunc
  PTG_FUNCVAR_V   constant pls_integer := 66;  -- 0x42 PtgFuncVar
  PTG_NAME_V      constant pls_integer := 67;  -- 0x43 PtgName
  PTG_REF_V       constant pls_integer := 68;  -- 0x44 PtgRef
  PTG_AREA_V      constant pls_integer := 69;  -- 0x45 PtgArea
  PTG_MEMAREA_V   constant pls_integer := 70;  -- 0x46 PtgMemArea
  PTG_MEMERR_V    constant pls_integer := 71;  -- 0x47 PtgMemErr
  PTG_MEMFUNC_V   constant pls_integer := 73;  -- 0x49 PtgMemFunc
  PTG_NAMEX_V     constant pls_integer := 89;  -- 0x59 PtgNameX
  PTG_REF3D_V     constant pls_integer := 90;  -- 0x5A PtgRef3d
  PTG_AREA3D_V    constant pls_integer := 91;  -- 0x5B PtgArea3d
  -- Array-type Ptg's
  PTG_ARRAY_A     constant pls_integer := 96;  -- 0x60 PtgArray
  PTG_FUNC_A      constant pls_integer := 97;  -- 0x61 PtgFunc
  PTG_FUNCVAR_A   constant pls_integer := 98;  -- 0x62 PtgFuncVar
  PTG_NAME_A      constant pls_integer := 99;  -- 0x63 PtgName
  PTG_REF_A       constant pls_integer := 100;  -- 0x64 PtgRef
  PTG_AREA_A      constant pls_integer := 101;  -- 0x65 PtgArea
  PTG_MEMAREA_A   constant pls_integer := 102;  -- 0x66 PtgMemArea
  PTG_MEMERR_A    constant pls_integer := 103;  -- 0x67 PtgMemErr
  PTG_MEMFUNC_A   constant pls_integer := 105;  -- 0x69 PtgMemFunc
  PTG_NAMEX_A     constant pls_integer := 121;  -- 0x79 PtgNameX
  PTG_REF3D_A     constant pls_integer := 122;  -- 0x7A PtgRef3d
  PTG_AREA3D_A    constant pls_integer := 123;  -- 0x7B PtgArea3d
  
  -- following Ptgs are extension to the official list, defined to assist parse and compile processes
  XPTG_PAREN_LEFT   constant pls_integer := 129;  -- 0x81 Extended Ptg : left parenthesis
  XPTG_PAREN_RIGHT  constant pls_integer := 130;  -- 0x82 Extended Ptg : right parenthesis
  XPTG_FUNC_STOP    constant pls_integer := 131;  -- 0x83 Extended Ptg : function stop
  XPTG_PARAM_SEP    constant pls_integer := 132;  -- 0x84 Extended Ptg : parameter separator
  XPTG_SPACE        constant pls_integer := 133;  -- 0x85 Extended Ptg : whitespace

  EXPR_UNARY         constant pls_integer := 256;  -- 0x0100 unary-expression
  EXPR_BINARY_REF    constant pls_integer := 257;  -- 0x0101 binary-reference-expression
  EXPR_BINARY_VAL    constant pls_integer := 258;  -- 0x0102 binary-value-expression
  EXPR_FUNC_CALL     constant pls_integer := 259;  -- 0x0103 function-call
  EXPR_DISPLAY_PREC  constant pls_integer := 260;  -- 0x0104 display-precedence-specifier
  EXPR_MEM_AREA      constant pls_integer := 261;  -- 0x0105 mem-area-expression
  
  MATCH_FUNC_START  constant pls_integer := 0;
  MATCH_FUNC_STOP   constant pls_integer := 1;

  RT_CELL         constant pls_integer := 3;
  RT_ROW          constant pls_integer := 1;
  RT_COLUMN       constant pls_integer := 2;
  RT_NAMED_RANGE  constant pls_integer := 4;
  
  OP_PLUS       constant binary_integer := 1; -- +
  OP_MINUS      constant binary_integer := 2; -- -
  OP_MUL        constant binary_integer := 3; -- *
  OP_DIV        constant binary_integer := 4; -- /
  OP_EXP        constant binary_integer := 5; -- ^
  OP_CONCAT     constant binary_integer := 6; -- &
  OP_LT         constant binary_integer := 7; -- < 
  OP_LE         constant binary_integer := 8; -- <= 
  OP_EQ         constant binary_integer := 9; -- =
  OP_GE         constant binary_integer := 10; -- >=
  OP_GT         constant binary_integer := 11; -- >
  OP_NE         constant binary_integer := 12; -- <>
  OP_UNION      constant binary_integer := 13; -- ,
  OP_RANGE      constant binary_integer := 14; -- :
  OP_PERCENT    constant binary_integer := 15; -- %
  
  type intList_t is table of pls_integer;
  tokenOpPtgMap  intList_t := intList_t(
  PTG_ADD, PTG_SUB, PTG_MUL, PTG_DIV, PTG_POWER, PTG_CONCAT, PTG_LT, PTG_LE, PTG_EQ, PTG_GE, PTG_GT, PTG_NE, PTG_UNION, PTG_RANGE, PTG_PERCENT);

  T_LEFT        constant binary_integer := 30; -- (
  T_RIGHT       constant binary_integer := 31; -- )
  T_COMMA       constant binary_integer := 32; -- ,
  
  T_SEMICOLON   constant binary_integer := 35; -- ;
  T_BANG        constant binary_integer := 36; -- ! 0x24

  T_NUMBER      constant binary_integer := 40;
  
  T_FUNC_START  constant binary_integer := 42;
  T_STRING      constant binary_integer := 45;
  
  T_ARRAY_START     constant binary_integer := 47;
  T_ARRAYROW_START  constant binary_integer := 48;
  T_OPERAND         constant binary_integer := 49;
  T_ARG_SEP         constant binary_integer := 50;
  T_MISSING_ARG     constant binary_integer := 51;
  
  T_ARRAYITEM_SEP  constant binary_integer := 52;
  T_ARRAYROW_STOP  constant binary_integer := 53;
  T_WSPACE         constant binary_integer := 54;
  T_ARRAY_STOP     constant binary_integer := 55;
  T_FUNC_STOP      constant binary_integer := 56;
  T_PREFIX         constant binary_integer := 57;
  T_ERROR          constant binary_integer := 58;
  T_QUOTED         constant binary_integer := 59; -- 0x3B
  T_BOOLEAN        constant binary_integer := 60;
  
  T_BASE_REF     constant binary_integer := 64; -- 0x40
  T_ROW_REF      constant binary_integer := T_BASE_REF + RT_ROW; -- 0x41
  T_COL_REF      constant binary_integer := T_BASE_REF + RT_COLUMN; -- 0x42
  T_CELL_REF     constant binary_integer := T_BASE_REF + RT_CELL; -- 0x43
  T_NAME         constant binary_integer := T_BASE_REF + RT_NAMED_RANGE; -- 0x44
  T_SHEET        constant binary_integer := 80; -- 0x50
  
  type tokenTypeList_t is table of pls_integer;
  function tokenTypeSequence (tokenTypes in tokenTypeList_t) return varchar2;
  
  -- sequences of tokens that parse as a single ptg:
  QUOTED_MULTISHEET_ROW_RANGE   constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_PREFIX, T_BANG, T_ROW_REF, OP_RANGE, T_ROW_REF));
  QUOTED_MULTISHEET_COL_RANGE   constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_PREFIX, T_BANG, T_COL_REF, OP_RANGE, T_COL_REF));
  QUOTED_MULTISHEET_CELL_RANGE  constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_PREFIX, T_BANG, T_CELL_REF, OP_RANGE, T_CELL_REF));
  QUOTED_MULTISHEET_CELL_REF    constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_PREFIX, T_BANG, T_CELL_REF));
  
  MULTISHEET_ROW_RANGE          constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, OP_RANGE, T_SHEET, T_BANG, T_ROW_REF, OP_RANGE, T_ROW_REF));
  MULTISHEET_COL_RANGE          constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, OP_RANGE, T_SHEET, T_BANG, T_COL_REF, OP_RANGE, T_COL_REF));
  MULTISHEET_CELL_RANGE         constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, OP_RANGE, T_SHEET, T_BANG, T_CELL_REF, OP_RANGE, T_CELL_REF));
  MULTISHEET_CELL_REF           constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, OP_RANGE, T_SHEET, T_BANG, T_CELL_REF));

  SHEET_ROW_RANGE               constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, T_BANG, T_ROW_REF, OP_RANGE, T_ROW_REF));
  SHEET_COL_RANGE               constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, T_BANG, T_COL_REF, OP_RANGE, T_COL_REF));
  SHEET_CELL_RANGE              constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, T_BANG, T_CELL_REF, OP_RANGE, T_CELL_REF));
  SHEET_CELL_REF                constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, T_BANG, T_CELL_REF));
  
  SCOPED_NAME                   constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, T_BANG, T_NAME));
  SHEET_PREFIX                  constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_SHEET, T_BANG));
  
  ROW_RANGE                     constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_ROW_REF, OP_RANGE, T_ROW_REF));
  COL_RANGE                     constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_COL_REF, OP_RANGE, T_COL_REF));
  CELL_RANGE                    constant varchar2(2048) := tokenTypeSequence(tokenTypeList_t(T_CELL_REF, OP_RANGE, T_CELL_REF));
    
  --LETTERS         constant varchar2(26) := 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  --DIGITS          constant varchar2(10) := '0123456789';
  MAX_COL_NUMBER  constant pls_integer := 16384;
  MAX_ROW_NUMBER  constant pls_integer := 1048576;
  
  decimalSep    varchar2(1);
  
  -- valueTypes
  VT_NONE       constant pls_integer := 0;
  VT_REFERENCE  constant pls_integer := 1;
  VT_VALUE      constant pls_integer := 2;
  VT_ARRAY      constant pls_integer := 3;
  
  /** Effective token class conversion types. */
  --enum XclExpClassConv
  EXC_CLASSCONV_ORG constant pls_integer := 0;          -- Keep original class of the token.
  EXC_CLASSCONV_VAL constant pls_integer := 1;          -- Convert ARR tokens to VAL class (REF remains unchanged).
  EXC_CLASSCONV_ARR constant pls_integer := 2;          -- Convert VAL tokens to ARR class (REF remains unchanged).
  
  /** Enumerates different types of token class conversion in function parameters. */
  --enum XclFuncParamConv
  EXC_PARAMCONV_ORG constant pls_integer := 0;          -- Use original class of current token.
  EXC_PARAMCONV_VAL constant pls_integer := 1;          -- Convert tokens to VAL class.
  EXC_PARAMCONV_ARR constant pls_integer := 2;          -- Convert tokens to ARR class.
  EXC_PARAMCONV_RPT constant pls_integer := 3;          -- Repeat parent conversion in VALTYPE parameters.
  EXC_PARAMCONV_RPX constant pls_integer := 4;          -- Repeat parent conversion in REFTYPE parameters.
  EXC_PARAMCONV_RPO constant pls_integer := 5;          -- Repeat parent conversion in operands of operators.

  /** Type of a formula. */
  --enum XclFormulaType
  --EXC_FMLATYPE_CELL     constant pls_integer := 0;   -- Simple cell formula, also used in change tracking.
  EXC_FMLATYPE_MATRIX   constant pls_integer := 1;   -- Matrix (array) formula.
  EXC_FMLATYPE_SHARED   constant pls_integer := 2;   -- Shared formula.
  EXC_FMLATYPE_CONDFMT  constant pls_integer := 3;   -- Conditional format.
  EXC_FMLATYPE_DATAVAL  constant pls_integer := 4;   -- Data validation.
  --EXC_FMLATYPE_NAME     constant pls_integer := 5;   -- Defined name.
  EXC_FMLATYPE_CHART    constant pls_integer := 6;   -- Chart source ranges.
  --EXC_FMLATYPE_CONTROL  constant pls_integer := 7;   -- Spreadsheet links in form controls.
  --EXC_FMLATYPE_WQUERY   constant pls_integer := 8;   -- Web query source range.
  EXC_FMLATYPE_LISTVAL  constant pls_integer := 9;   -- List (cell range) validation.

  /** Type of token class handling. */
  --enum XclExpFmlaClassType
  EXC_CLASSTYPE_CELL    constant pls_integer := 0;   -- Cell formula, shared formula.
  EXC_CLASSTYPE_ARRAY   constant pls_integer := 1;   -- Array formula, conditional formatting, data validation.
  EXC_CLASSTYPE_NAME    constant pls_integer := 2;   -- Defined name, range list.
  
  type operandInfo_t is record (
    convClass  pls_integer -- operand conversion class
  , valueType  boolean -- expected type (true=Value/false=Reference) when this token is an operand of function or operator
  );
  
  type operandInfoList_t is table of operandInfo_t;

  type functionMetadata_t is record (
    id             pls_integer
  , name           varchar2(128)
  , minParamCount  pls_integer
  , maxParamCount  pls_integer
  , returnClass    pls_integer
  , paramInfos     operandInfoList_t
  , volatile       boolean
  , internalName   varchar2(128)
  );
  
  type functionMetadataMap_t is table of functionMetadata_t index by varchar2(128);
  
  type sheet_t is record (idx pls_integer, name sheet_name, quotable boolean := false);
  type sheetMap_t is table of sheet_t index by varchar2(128);
  
  type sheetRange_t is record (firstSheet sheet_t, lastSheet sheet_t, isSingleSheet boolean := true);
  type worksheetPrefix_t is record (workbookIdx pls_integer, sheetRange sheetRange_t, isQuoted boolean := false, isNull boolean := true);
  
  type row_t is record (value pls_integer, isAbsolute boolean := false);
  type column_t is record (value pls_integer, alphaValue varchar2(3), isAbsolute boolean := false);
  type cell_t is record (col column_t, rw row_t, type pls_integer := 0, isNull boolean := true);
  type area_t is record (firstCell cell_t, lastCell cell_t);
  type area3d_t is record (prefix worksheetPrefix_t, area area_t);
  
  type areaList_t is table of area_t;
  
  NULL_CELL    cell_t;
  NULL_AREA    area_t;
  NULL_AREA3D  area3d_t;
  
  -- defined name: the scope should match a sheet name, or null for a workbook-level name
  type name_t is record (idx pls_integer, value varchar2(255 char), scope sheet_t);
  type nameMap_t is table of name_t index by varchar2(2048);
    
  type arrayItem_t is record (itemType pls_integer, strValue varchar2(255 char), errValue pls_integer, boolValue pls_integer, numValue binary_double);
  type arrayRow_t is table of arrayItem_t;
  type array_t is table of arrayRow_t;

  type pos_t is record (
    ln       pls_integer -- line number
  , cn       pls_integer -- column number
  --, last_cn  pls_integer -- last column number of previous line
  );
  
  type parseNode_t is record (
    id         pls_integer
  , nodeType   pls_integer
  , attrType   pls_integer
  , children   intList_t
  , parentId   pls_integer
  , token      varchar2(32767)
  , pos        pos_t
  , nodeClass  pls_integer
  , operandInfo  operandInfo_t
  );
  
  type parseNodeList_t is table of parseNode_t;
  --type parseNodeMap_t is table of parseNode_t index by pls_integer;
  
  type ptgBool_t is record (value binary_integer);
  type ptgNum_t is record (value binary_double);
  type ptgInt_t is record (value binary_integer);
  type ptgStr_t is record (value varchar2(255 char));
  type ptgErr_t is record (value pls_integer);
  type ptgArray_t is record (value array_t);
  type ptgArea3d_t is record (value area3d_t);
  type ptgName_t is record (value name_t);
  type ptgFuncVar_t is record (name varchar2(256), argc pls_integer);
  type ptgMemErr_t is record (value pls_integer);
  type ptgMemArea_t is record (value areaList_t);
  
  type ptgBoolList_t is table of ptgBool_t index by pls_integer;
  type ptgNumList_t is table of ptgNum_t index by pls_integer;
  type ptgIntList_t is table of ptgInt_t index by pls_integer;
  type ptgStrList_t is table of ptgStr_t index by pls_integer;
  type ptgErrList_t is table of ptgErr_t index by pls_integer;
  type ptgArrayList_t is table of ptgArray_t index by pls_integer;
  type ptgArea3dList_t is table of ptgArea3d_t index by pls_integer;
  type ptgNameList_t is table of ptgName_t index by pls_integer;
  type ptgFuncVarList_t is table of ptgFuncVar_t index by pls_integer;
  type ptgMemErrList_t is table of ptgMemErr_t index by pls_integer;
  type ptgMemAreaList_t is table of ptgMemArea_t index by pls_integer;
  
  parse_exception  exception;
  pragma exception_init(parse_exception, -20000);
  
  functionMetadataMap  functionMetadataMap_t;
  
  type stringList_t is table of varchar2(4000);
  
  type labelMap_t is table of varchar2(256) index by pls_integer;
  tokenLabels  labelMap_t;
  ptgLabels    labelMap_t;
  
  type identifier_t is record (value varchar2(255 char), matchSheetName boolean := false, matchDefinedName boolean := false);
  type operand_t is record (prefix worksheetPrefix_t, cell cell_t, ident identifier_t);
  
  type token_t is record (value varchar2(256), type pls_integer, parsedValue operand_t, pos pos_t);
  type tokenList_t is table of token_t;
  
  type tokenStream_t is record (tokens tokenList_t := tokenList_t(), idx pls_integer := 0);
  
  type errorMap_t is table of pls_integer index by varchar2(16);
  
  type opMetadata_t is record (
    token        varchar2(2)
  , argc         pls_integer
  , prec         pls_integer
  , assoc        pls_integer
  , returnClass  pls_integer
  , operandInfo  operandInfo_t
  );
  
  type opMetadataMap_t is table of opMetadata_t index by pls_integer;
  
  --type quotableCharMap_t is table of pls_integer index by varchar2(1 char);

  type config_t is record (
    fmlaType       pls_integer  -- EXC_FMLATYPE_*
  , fmlaClassType  pls_integer  -- EXC_CLASSTYPE_*
  );
  
  type byteArray_t is record (content raw(32767), lobContent blob);
  type charArray_t is record (content varchar2(32767), lobContent clob);
  
  type stream_t is record (
    chars  charArray_t
  , bytes  byteArray_t
  , sz     pls_integer := 0
  , isLob  boolean := false
  );

  type context_t is record (
    sheet      sheet_t
  , cell       cell_t
  , sheetMap   sheetMap_t
  , nameMap    nameMap_t
  , externals  ExcelTypes.CT_Externals
  , config     config_t
  , binary     boolean
  , volatile   boolean
  , rgce       stream_t
  , rgbExtra   stream_t
  -- Defined names generated by this formula, e.g. when a future function is used (_xlfn.XXX)
  -- This collection is meant to be sent back to the calling context for merging, and deleted
  , definedNames  ExcelTypes.CT_DefinedNames
  , treeRootId  pls_integer
  );
  
  ctx  context_t; 
  
  opMetadataMap   opMetadataMap_t;
  errorMap        errorMap_t;
  --quotableCharMap  quotableCharMap_t;
  
  formulaString   varchar2(32767);
  formulaLength   pls_integer;
  pointer         pls_integer := 0;
  look            varchar2(1 char);
  pos             pos_t;
  
  t               parseNodeList_t := parseNodeList_t();
  
  ptgBoolList     ptgBoolList_t;
  ptgNumList      ptgNumList_t;
  ptgIntList      ptgIntList_t;
  ptgStrList      ptgStrList_t;
  ptgErrList      ptgErrList_t;
  ptgArrayList    ptgArrayList_t;
  ptgArea3dList   ptgArea3dList_t;
  ptgNameList     ptgNameList_t;
  ptgFuncVarList  ptgFuncVarList_t;
  ptgMemErrList   ptgMemErrList_t;
  ptgMemAreaList  ptgMemAreaList_t;
  
  --streams         streamStack_t;
  --parentStream    stream_t;
  currentStream   stream_t;

  procedure putChars (stream in out nocopy stream_t, chars in varchar2) is
    len  pls_integer := length(chars);
  begin
    if stream.isLob then
      dbms_lob.writeappend(stream.chars.lobContent, len, chars);
      stream.sz := stream.sz + len;
    else
      if lengthb(stream.chars.content) + lengthb(chars) > 32767 then
        -- switch to lob storage
        dbms_lob.createtemporary(stream.chars.lobContent, true);
        stream.isLob := true;
        dbms_lob.writeappend(stream.chars.lobContent, stream.sz, stream.chars.content);
        dbms_lob.writeappend(stream.chars.lobContent, len, chars);
        stream.sz := stream.sz + len;
      else
        stream.chars.content := stream.chars.content || chars;
        stream.sz := stream.sz + len;
      end if;
    end if;
  end;

  -- append stream2 to stream1
  procedure putChars (stream1 in out nocopy stream_t, stream2 in stream_t) is
  begin
    if stream2.isLob then
      for i in 0 .. ceil(stream2.sz/8191) - 1 loop
        putChars(stream1, dbms_lob.substr(stream2.chars.lobContent, 8191, 8191*i + 1));
      end loop;
      --dbms_lob.freetemporary(stream2.chars.lobContent);
    else
      putChars(stream1, stream2.chars.content);
    end if;
  end;

  procedure putBytes (stream in out nocopy stream_t, bytes in raw, offset in pls_integer default null) is
    len  pls_integer := utl_raw.length(bytes);
  begin
    -- offset parameter MUST be less than or equal to ( stream.sz - sizeof(bytes) + 1 )
    if stream.isLob then
      dbms_lob.write(stream.bytes.lobContent, len, nvl(offset, stream.sz + 1), bytes);
      stream.sz := stream.sz + len;
    else
      if stream.sz + len > 32767 then
        -- switch to lob storage
        dbms_lob.createtemporary(stream.bytes.lobContent, true);
        stream.isLob := true;
        dbms_lob.writeappend(stream.bytes.lobContent, stream.sz, stream.bytes.content);
        dbms_lob.writeappend(stream.bytes.lobContent, len, bytes);
        stream.sz := stream.sz + len;
      else
        if stream.bytes.content is null then
          stream.bytes.content := bytes;
        else
          stream.bytes.content := utl_raw.overlay(bytes, stream.bytes.content, nvl(offset, stream.sz + 1));
        end if;
        stream.sz := stream.sz + len;
      end if;
    end if;
  end;
  
  procedure initStream (stream in out nocopy stream_t) is
  begin
    stream.bytes.content := null;
    stream.sz := 0;
    stream.isLob := false;
    if dbms_lob.istemporary(stream.bytes.lobContent) = 1 then
      dbms_lob.freetemporary(stream.bytes.lobContent);
    end if;
  end;
  
  procedure put (chars in varchar2) is
  begin
    putChars(currentStream, chars);
  end;
  
  procedure putBytes (bytes in raw, offset in pls_integer default null) is
  begin
    putBytes(currentStream, bytes, offset);
  end;
  
  procedure putExtra (bytes in raw) is
  begin
    putBytes(ctx.rgbExtra, bytes);
  end;  

  function parseStringList (
    input  in varchar2
  , sep    in varchar2 
  )
  return stringList_t
  is
    token   varchar2(4000);
    p1      pls_integer := 1;
    p2      pls_integer;
    output  stringList_t := stringList_t();
  begin
    if input is not null then
      loop
        p2 := instr(input, sep, p1);
        if p2 = 0 then
          token := substr(input, p1);
        else
          token := substr(input, p1, p2-p1);    
          p1 := p2 + 1;
        end if;
        output.extend;
        output(output.last) := token;
        exit when p2 = 0;
      end loop;
    end if;
    return output;   
  end;

  function makeOperandInfo (
    operandInfoStr in varchar2
  )
  return operandInfo_t
  is
    info  operandInfo_t;
  begin
    info.convClass := case substr(operandInfoStr, 2, 1)
                      when 'O' then EXC_PARAMCONV_ORG
                      when 'A' then EXC_PARAMCONV_ARR
                      when 'R' then EXC_PARAMCONV_RPT
                      when 'X' then EXC_PARAMCONV_RPX
                      when 'V' then EXC_PARAMCONV_VAL
                      when 'P' then EXC_PARAMCONV_RPO
                      end;
    info.valueType := ( substr(operandInfoStr, 1, 1) = 'V' );
    return info;
  end;

  function makeOpMetadata (
    token           in varchar2
  , argc            in pls_integer
  , prec            in pls_integer
  , assoc           in pls_integer default 0
  , returnClass     in pls_integer default VT_VALUE
  , operandInfoStr  in varchar2 default 'VP'
  )
  return opMetadata_t
  is
    meta  opMetadata_t;
  begin
    meta.token := token;
    meta.argc := argc;
    meta.prec := prec;
    meta.assoc := assoc;
    meta.returnClass := returnClass;
    meta.operandInfo := makeOperandInfo(operandInfoStr);
    return meta;
  end;

  procedure initOpMetadata is
  begin
    -- lexical value, argument count, precedence (higher is greater), associativity (0=left, 1=right), operator class
    opMetadataMap(PTG_LT)      := makeOpMetadata('<', 2, 1);
    opMetadataMap(PTG_LE)      := makeOpMetadata('<=', 2, 1);
    opMetadataMap(PTG_EQ)      := makeOpMetadata('=', 2, 1);
    opMetadataMap(PTG_GE)      := makeOpMetadata('>=', 2, 1);
    opMetadataMap(PTG_GT)      := makeOpMetadata('>', 2, 1);
    opMetadataMap(PTG_NE)      := makeOpMetadata('<>', 2, 1);
    opMetadataMap(PTG_CONCAT)  := makeOpMetadata('&', 2, 2);    
    opMetadataMap(PTG_ADD)     := makeOpMetadata('+', 2, 3);
    opMetadataMap(PTG_SUB)     := makeOpMetadata('-', 2, 3);    
    opMetadataMap(PTG_MUL)     := makeOpMetadata('*', 2, 4);
    opMetadataMap(PTG_DIV)     := makeOpMetadata('/', 2, 4);
    opMetadataMap(PTG_POWER)   := makeOpMetadata('^', 2, 5);
    opMetadataMap(PTG_PERCENT) := makeOpMetadata('%', 1, 6);
    opMetadataMap(PTG_UMINUS)  := makeOpMetadata('-', 1, 7);
    opMetadataMap(PTG_UNION)   := makeOpMetadata(',', 2, 8, returnClass => VT_REFERENCE, operandInfoStr => 'RP');
    opMetadataMap(PTG_ISECT)   := makeOpMetadata(' ', 2, 9, returnClass => VT_REFERENCE, operandInfoStr => 'RP');
    opMetadataMap(PTG_RANGE)   := makeOpMetadata(':', 2, 10, returnClass => VT_REFERENCE, operandInfoStr => 'RP');
  end;

  procedure initFuncMetadataMap
  is
    strings  stringList_t;
    meta     functionMetadata_t;
    
    cursor c_meta is
    select funcname
         , xclfunc
         , minparamcount
         , maxparamcount
         , retclass
         , paraminfos
         , case when flags like '%VOLATILE%' then 1 else 0 end as volatile
         , xclfuncname
    from xclfunctioninfo
    order by funcname;
    
  begin
    
    for r in c_meta loop

      meta.id := r.xclfunc;
      meta.name := r.funcname;
      meta.minParamCount := r.minparamcount;
      meta.maxParamCount := r.maxparamcount;
      meta.returnClass := case r.retclass 
                          when 'R' then VT_REFERENCE
                          when 'V' then VT_VALUE
                          when 'A' then VT_ARRAY
                          end;
      meta.volatile := ( r.volatile = 1 );
      meta.internalName := r.xclfuncname;
      
      strings := parseStringList(r.paraminfos, '|');
      meta.paramInfos := operandInfoList_t();
      for i in 1 .. strings.count loop
        meta.paramInfos.extend;
        meta.paramInfos(i) := makeOperandInfo(strings(i));
      end loop;
      
      functionMetadataMap(meta.name) := meta;
    
    end loop;
       
  end;

  procedure initErrorMap is
  begin
    errorMap('#NULL!')        := 0;  -- 0x00
    errorMap('#DIV/0!')       := 7;  -- 0x07
    errorMap('#VALUE!')       := 15;  -- 0x0F
    errorMap('#REF!')         := 23;  -- 0x17
    errorMap('#NAME?')        := 29;  -- 0x1D
    errorMap('#NUM!')         := 36;  -- 0x24
    errorMap('#N/A')          := 42;  -- 0x2A
    errorMap('#GETTING_DATA') := 43;  -- 0x2B
  end;

  procedure initLabels is
  begin

    tokenLabels(OP_MINUS) := '-';
    tokenLabels(OP_PLUS) := '+';
    tokenLabels(OP_MUL) := '*';
    tokenLabels(OP_DIV) := '/';
    tokenLabels(OP_EXP) := '^';
    tokenLabels(OP_EQ) := '=';
    tokenLabels(OP_LT) := '<';
    tokenLabels(OP_GT) := '>';
    tokenLabels(OP_LE) := '<=';
    tokenLabels(OP_GE) := '>=';
    tokenLabels(OP_NE) := '<>';
    tokenLabels(OP_CONCAT) := '&';
    tokenLabels(OP_PERCENT) := '%';
       
    tokenLabels(T_LEFT) := '(';
    tokenLabels(T_RIGHT) := ')';
    tokenLabels(T_COMMA) := ',';
    tokenLabels(OP_RANGE) := 'range-operator';
    tokenLabels(T_SEMICOLON) := ';';

    tokenLabels(T_NUMBER) := 'number';
    tokenLabels(T_FUNC_START) := 'function-start';
    tokenLabels(T_STRING) := 'string';
    tokenLabels(T_ARRAY_START) := 'array-start';
    tokenLabels(T_ARRAYROW_START) := 'array-row-start';
    tokenLabels(T_OPERAND) := 'operand';
    tokenLabels(T_ARG_SEP) := 'argument-separator';
    tokenLabels(T_MISSING_ARG) := 'missing-argument';
    tokenLabels(OP_UNION) := 'union';
    tokenLabels(T_ARRAYITEM_SEP) := 'array-item-separator';
    tokenLabels(T_ARRAYROW_STOP) := 'array-row-stop';
    tokenLabels(T_WSPACE) := 'whitespace';
    tokenLabels(T_ARRAY_STOP) := 'array-stop';
    tokenLabels(T_FUNC_STOP) := 'function-stop';
    tokenLabels(T_PREFIX) := 'work-sheet-prefix';
    tokenLabels(T_BANG) := '!';
    tokenLabels(T_ERROR) := 'error';
    tokenLabels(T_QUOTED) := 'quoted-name';
    tokenLabels(T_ROW_REF) := 'single-row-reference';
    tokenLabels(T_COL_REF) := 'single-column-reference';
    tokenLabels(T_CELL_REF) := 'single-cell-reference';
    tokenLabels(T_NAME) := 'name';
    --tokenLabels(T_SHEET) := 'sheet-name';
    tokenLabels(T_BOOLEAN) := 'boolean';

    ptgLabels(PTG_ADD)          := 'PtgAdd';
    ptgLabels(PTG_SUB)          := 'PtgSub';
    ptgLabels(PTG_MUL)          := 'PtgMul';
    ptgLabels(PTG_DIV)          := 'PtgDiv';
    ptgLabels(PTG_POWER)        := 'PtgPower';
    ptgLabels(PTG_CONCAT)       := 'PtgConcat';
    ptgLabels(PTG_LT)           := 'PtgLt';
    ptgLabels(PTG_LE)           := 'PtgLe';
    ptgLabels(PTG_EQ)           := 'PtgEq';
    ptgLabels(PTG_GE)           := 'PtgGe';
    ptgLabels(PTG_GT)           := 'PtgGt';
    ptgLabels(PTG_NE)           := 'PtgNe';
    ptgLabels(PTG_ISECT)        := 'PtgIsect';
    ptgLabels(PTG_UNION)        := 'PtgUnion';
    ptgLabels(PTG_RANGE)        := 'PtgRange';
    ptgLabels(PTG_UPLUS)        := 'PtgUPlus';
    ptgLabels(PTG_UMINUS)       := 'PtgUMinus';
    ptgLabels(PTG_PERCENT)      := 'PtgPercent';
    ptgLabels(PTG_PAREN)        := 'PtgParen';
    ptgLabels(PTG_MISSARG)      := 'PtgMissArg';
    ptgLabels(PTG_STR)          := 'PtgString';
    ptgLabels(PTG_ERR)          := 'PtgErr';
    ptgLabels(PTG_BOOL)         := 'PtgBool';
    ptgLabels(PTG_INT)          := 'PtgInt';
    ptgLabels(PTG_NUM)          := 'PtgNum';
    
    ptgLabels(PTG_ARRAY_R)      := 'PtgArray';
    ptgLabels(PTG_ARRAY_V)      := 'PtgArray';
    ptgLabels(PTG_ARRAY_A)      := 'PtgArray';
    ptgLabels(PTG_FUNC_R)       := 'PtgFunc';
    ptgLabels(PTG_FUNC_V)       := 'PtgFunc';
    ptgLabels(PTG_FUNC_A)       := 'PtgFunc';
    ptgLabels(PTG_FUNCVAR_R)    := 'PtgFuncVar';
    ptgLabels(PTG_FUNCVAR_V)    := 'PtgFuncVar';
    ptgLabels(PTG_FUNCVAR_A)    := 'PtgFuncVar';
    ptgLabels(PTG_NAME_R)       := 'PtgName';
    ptgLabels(PTG_NAME_V)       := 'PtgName';
    ptgLabels(PTG_NAME_A)       := 'PtgName';
    ptgLabels(PTG_REF_R)        := 'PtgRef';
    ptgLabels(PTG_REF_V)        := 'PtgRef';
    ptgLabels(PTG_REF_A)        := 'PtgRef';
    ptgLabels(PTG_AREA_R)       := 'PtgArea';
    ptgLabels(PTG_AREA_V)       := 'PtgArea';
    ptgLabels(PTG_AREA_A)       := 'PtgArea';
    ptgLabels(PTG_MEMAREA_R)    := 'PtgMemArea';
    ptgLabels(PTG_MEMAREA_V)    := 'PtgMemArea';
    ptgLabels(PTG_MEMAREA_A)    := 'PtgMemArea';
    ptgLabels(PTG_MEMERR_R)     := 'PtgMemErr';
    ptgLabels(PTG_MEMERR_V)     := 'PtgMemErr';
    ptgLabels(PTG_MEMERR_A)     := 'PtgMemErr';
    ptgLabels(PTG_MEMFUNC_R)    := 'PtgMemFunc';
    ptgLabels(PTG_MEMFUNC_V)    := 'PtgMemFunc';
    ptgLabels(PTG_MEMFUNC_A)    := 'PtgMemFunc';
    ptgLabels(PTG_NAMEX_R)      := 'PtgNameX';
    ptgLabels(PTG_NAMEX_V)      := 'PtgNameX';
    ptgLabels(PTG_NAMEX_A)      := 'PtgNameX';
    ptgLabels(PTG_REF3D_R)      := 'PtgRef3d';
    ptgLabels(PTG_REF3D_V)      := 'PtgRef3d';
    ptgLabels(PTG_REF3D_A)      := 'PtgRef3d';
    ptgLabels(PTG_AREA3D_R)     := 'PtgArea3d';
    ptgLabels(PTG_AREA3D_V)     := 'PtgArea3d';
    ptgLabels(PTG_AREA3D_A)     := 'PtgArea3d';
    
    ptgLabels(XPTG_PAREN_LEFT)  := 'ExtPtgParenLeft';
    ptgLabels(XPTG_PAREN_RIGHT) := 'ExtPtgParenRight';
    ptgLabels(XPTG_FUNC_STOP)   := 'ExtPtgFuncStop';
    ptgLabels(XPTG_PARAM_SEP)   := 'ExtPtgParamSep';
    ptgLabels(XPTG_SPACE)       := 'ExtPtgWhitespace';

    ptgLabels(EXPR_UNARY)        := 'unary-expression';
    ptgLabels(EXPR_BINARY_REF)   := 'binary-reference-expression';
    ptgLabels(EXPR_BINARY_VAL)   := 'binary-value-expression';
    ptgLabels(EXPR_FUNC_CALL)    := 'function-call';
    ptgLabels(EXPR_DISPLAY_PREC) := 'display-precedence-specifier';
    ptgLabels(EXPR_MEM_AREA)     := 'mem-area-expression';

  end;

  function ptgLabel(node in parseNode_t) return varchar2 is
  begin
    if node.nodeType = PTG_ATTR_B and node.attrType = PTG_ATTRSUM then
      return 'PtgAttrSum';
    else
      return ptgLabels(node.nodeType);
    end if;
  end;

  procedure applyPtgDataType (
    ptgType   in out nocopy pls_integer
  , dataType  in pls_integer
  ) 
  is
  begin
    ptgType := bitand(ptgType, 159) -- 10011111 (clear bits 5-6)
             + dataType * 32 ; -- 00100000
  end; 
  
  function newToken (tokenType in pls_integer, tokenValue in varchar2 default null, tokenPos in pos_t default null) return token_t is
    token  token_t;
  begin
    token.value := tokenValue;
    token.type := tokenType;
    token.pos := tokenPos;
    return token;    
  end;
  
  procedure pushToken (stream in out nocopy tokenStream_t, token in token_t)
  is
  begin
    stream.tokens.extend;
    stream.idx := stream.idx + 1;
    stream.tokens(stream.idx) := token;
  end;

  procedure pushToken (stream in out nocopy tokenStream_t, tokenValue in varchar2, tokenType in pls_integer, tokenPos in pos_t default null)
  is
  begin
    pushToken(stream, newToken(tokenType, tokenValue, tokenPos));
  end;
  
  function pushToken (stream in out nocopy tokenStream_t, tokenValue in varchar2, tokenType in pls_integer, tokenPos in pos_t default null) return token_t is
    token  token_t := newToken(tokenType, tokenValue, tokenPos);
  begin
    pushToken(stream, token);
    return token;
  end;
  
  function peekToken (stream in tokenStream_t, pos in pls_integer default 1) return token_t is
  begin
    if stream.idx >= nvl(pos, 1) then
      return stream.tokens(stream.idx - pos + 1);
    else
      return null;
    end if;
  end;
  
  procedure popToken (stream in out nocopy tokenStream_t) is
  begin
    if stream.idx != 0 then
      stream.tokens.trim;
      stream.idx := stream.idx - 1;
    end if;
  end;
  
  procedure append (token in out nocopy token_t, str in varchar2) is
  begin
    if token.value is null then
      -- save start position of this token
      token.pos := pos;
    end if;
    token.value := token.value || str;
  end;
  
  function parseCellReference (input in varchar2) return cell_t;
  
  procedure parseToken (token in out nocopy token_t, tokenType in pls_integer) is 
  begin
    --if token.value is not null then
      
      if tokenType = T_OPERAND then
        
        -- boolean?
        if token.value in ('TRUE', 'FALSE') then
          token.type := T_BOOLEAN;
        else
          -- try to parse a cell reference or name out of it
          token.parsedValue.cell := parseCellReference(token.value);
          if not token.parsedValue.cell.isNull then
            token.type := T_BASE_REF + token.parsedValue.cell.type;
          else
            token.parsedValue.ident.value := token.value;
            token.parsedValue.ident.matchSheetName := ctx.sheetMap.exists(upper(token.value));
            --TODO: check if the value matches a sheet name and/or a defined name
            --token.parsedValue.ident.matchDefinedName := nameMap.exists(token.value);
            token.type := T_NAME;
          end if;
        
        end if;
      
      else
        token.type := tokenType;
      end if;
      
    --end if;
  end;
  
  procedure pushAndClear (stream in out nocopy tokenStream_t, token in out nocopy token_t, tokenType in pls_integer) is
  begin
    if token.value is not null then
      parseToken(token, tokenType);
      pushToken(stream, token);
      token := null;
    end if;
  end;

  function pushAndClear (stream in out nocopy tokenStream_t, token in out nocopy token_t, tokenType in pls_integer)
  return token_t 
  is
    output  token_t;
  begin
    if token.value is not null then
      parseToken(token, tokenType);
      pushToken(stream, token);
      output := token;
      token := null;
    end if;
    return output;
  end;

  function nextToken (stream in out nocopy tokenStream_t, pos in pls_integer default 1) return token_t is
  begin
    if stream.idx <= stream.tokens.count - nvl(pos, 1) then
      return stream.tokens(stream.idx + nvl(pos, 1));
    else
      return null;
    end if;
  end;
  
  function previousNonWsToken (stream in out nocopy tokenStream_t) return token_t is
    pos    pls_integer := 1;
    token  token_t := peekToken(stream, pos);
  begin
    -- the tokenizer collapses whitespaces so there should be no sequence longer than one, but we'll loop anyway
    while token.type = T_WSPACE loop
      pos := pos + 1;
      token := peekToken(stream, pos);
    end loop;
    return token;
  end;

  function previousToken (stream in out nocopy tokenStream_t, pos in pls_integer default 1) return token_t is
  begin
    if stream.idx > nvl(pos, 1) then
      return stream.tokens(stream.idx - nvl(pos, 1));
    else
      return null;
    end if;
  end;
  
  function hasNextToken (stream in out nocopy tokenStream_t) return boolean is
  begin
    return ( stream.idx < stream.tokens.count );
  end;
  
  function getNextToken (stream in out nocopy tokenStream_t) return token_t is
  begin
    stream.idx := stream.idx + 1;
    return stream.tokens(stream.idx);
  end;
  
  function tokenTypeSequence (tokenTypes in tokenTypeList_t) return varchar2 is
    output  varchar2(2048);
  begin
    for i in 1 .. tokenTypes.count loop
      output := output || to_char(tokenTypes(i), 'FM0X');
    end loop;
    return output;
  end;

  function getTokenTypeSequence (stream in tokenStream_t, cnt in pls_integer) return varchar2 is
    offset     pls_integer := 0;
    maxOffset  pls_integer := least(cnt, stream.tokens.count - stream.idx + 1);
    token      token_t;
    output     varchar2(2048);
  begin
    while offset < maxOffset loop
      token := stream.tokens(stream.idx + offset);
      if token.type = T_NAME and token.parsedValue.ident.matchSheetName then
        output := output || to_char(T_SHEET, 'FM0X');
      else
        output := output || to_char(token.type, 'FM0X');
      end if;
      offset := offset + 1;
    end loop;
    return output;
  end;

  procedure error (
    p_message in varchar2
  , p_arg1 in varchar2 default null
  , p_arg2 in varchar2 default null
  , p_arg3 in varchar2 default null
  , p_arg4 in varchar2 default null
  )
  is
  begin
    raise_application_error(-20000, utl_lms.format_message(p_message, p_arg1, p_arg2, p_arg3, p_arg4));
  end;

  function getNLS (parameterName in varchar2) 
  return varchar2 
  is
    result  nls_session_parameters.value%type;
  begin
    select value
    into result
    from nls_session_parameters
    where parameter = parameterName ;
    return result;
  exception
    when no_data_found then
      return null;
  end;

  procedure initState is
  begin
    decimalSep := substr(getNLS('NLS_NUMERIC_CHARACTERS'), 1, 1);
    initOpMetadata;
    initFuncMetadataMap;
    initLabels;
    initErrorMap;
    --initQuotableCharMap;
  end;
  
  procedure addChildNode (id in pls_integer, childId in pls_integer) is
  begin
    t(id).children.extend;
    t(id).children(t(id).children.last) := childId;
  end;
  
  -- check whether character code point is greater than 0x7F (upper bound of ASCII range)
  function isUnicodeChar(c in varchar2) return boolean is
  begin
    return ( nlssort(c,'NLS_SORT=UNICODE_BINARY') > hextoraw('007F') );
  end;
  
  function isWhitespace(c in varchar2) return boolean is
  begin
    return ( c = ' ' or c = chr(10) or c = chr(13) or c = chr(9) );
  end;

  function isDigit(c in varchar2) return boolean is
  begin
    return ( c between '0' and '9' );
  end;

  function isLetter(c in varchar2) return boolean is
  begin
    return ( c between 'A' and 'Z' or c between 'a' and 'z' );
  end;
  
  function isAlpha (c in varchar2) return boolean is
  begin
    return ( isLetter(c) or c = '$' or c = '_' );
  end;

  function getSheet (sheetName in varchar2) return sheet_t is
  begin
    return ctx.sheetMap(upper(sheetName));
  end;
  
  function putName (value in varchar2, sheetName in varchar2 default null, idx in pls_integer default null)
  return name_t
  is
    nm       name_t;
    nameKey  varchar2(2048);
  begin
    nm.idx := nvl(idx, ctx.nameMap.count + 1);
    nm.value := value;
    if sheetName is not null then
      -- should match a declared sheet name
      if ctx.sheetMap.exists(upper(sheetName)) then
        nm.scope := getSheet(sheetName);
      else
        error('Invalid sheet name: %s', sheetName);
      end if;
    end if;
    nameKey := upper(case when nm.scope.idx is not null then nm.scope.name || '!' end || nm.value); 
    ctx.nameMap(nameKey) := nm;
    return nm;
  end;

  procedure putName (value in varchar2, sheetName in varchar2 default null, idx in pls_integer default null) is
    nm  name_t;
  begin
    nm := putName(value, sheetName, idx);
  end;

  procedure putSheet (name in varchar2, idx in pls_integer default null) is
    sheet  sheet_t;
  begin
    sheet.idx := nvl(idx, ctx.sheetMap.count + 1);
    sheet.name := name;
    sheet.quotable := false;
    -- check if sheet name contains a quotable character or matches a reserved sequence
    if isValidCellReference(upper(sheet.name)) then
      -- name matches an A1-reference
      sheet.quotable := true;
      --TODO: R1C1 reference
    else
      -- search for a quotable character
      if ExcelTypes.isSheetQuotableStartChar(substr(sheet.name, 1, 1)) then
        sheet.quotable := true;
      else
        
        for j in 2 .. length(sheet.name) loop
          if ExcelTypes.isSheetQuotableChar(substr(sheet.name, j, 1)) then
            sheet.quotable := true;
            exit;
          end if;
        end loop; 
               
      end if;
    end if;
    -- case-insensitive key
    ctx.sheetMap(upper(sheet.name)) := sheet;
  end;

  procedure setCurrentSheet (sheetName in varchar2) is
  begin
    ctx.sheet := ctx.sheetMap(upper(sheetName));
  end;

  procedure setCurrentCell (cellRef in varchar2) is
  begin
    ctx.cell := parseCellReference(cellRef);
  end;

  procedure setFormulaType (fmlaType in pls_integer) is
  begin
    ctx.config.fmlaType := fmlaType;
  end;
  
  function getPrevNonDigit return varchar2 is
  begin
    if pointer > 1 then
      for i in reverse 1 .. pointer - 1 loop
        if not isDigit(substr(formulaString, i, 1)) then
          return substr(formulaString, i, 1);
        end if;
      end loop;
    end if;
    return null;
  end;

  function getNextNonDigit return varchar2 is
  begin
    if pointer < formulaLength then
      for i in pointer + 1 .. formulaLength loop
        if not isDigit(substr(formulaString, i, 1)) then
          return substr(formulaString, i, 1);
        end if;
      end loop;
    end if;
    return null;
  end;

  function isPrevNonDigitBlank return boolean is
  begin
    return ( nvl(getPrevNonDigit, ' ') = ' ' );
  end;
  
  function isPrevOrNextNonDigitRangeOp return boolean is
  begin
    return ( nvl(getPrevNonDigit, ' ') = ':' or nvl(getNextNonDigit, ' ') = ':' );
  end;

  function toLocalNumber (str in varchar2) return number is
  begin
    return to_number(replace(str, '.', decimalSep));
  end;

  function toDouble (str in varchar2) return binary_double is
  begin
    return to_binary_double(replace(str, '.', decimalSep));
  end;
   
  -- sheet-name-character
  /*
  function isSheetNameChar (c in varchar2) return boolean is
  begin
    return c not in (':', ',', ' ', '^', '*', '/', '+', '-', '&', '=', '<', '>', '%', -- operators
                     '''', '[', ']', '\', '?', 
                     '(', ')', ';', '{', '}', '#', '"', '!');
  end;  
  */
  
  function eof return boolean is
  begin
    return ( pointer > formulaLength );
  end;
  
  procedure getChar is
  begin
    
    if look = chr(10) then
      pos.ln := pos.ln + 1;
      pos.cn := 0;
    end if;
    
    pointer := pointer + 1;
    pos.cn := pos.cn + 1;
    
    look := substr(formulaString, pointer, 1);   
    
  end;
  
  function nextChar return varchar2 is
  begin
    return substr(formulaString, pointer + 1, 1);
  end;

  procedure skipWhitespace is
  begin
    while isWhitespace(look) loop
      getChar;
    end loop;
  end;

  function isNull (area in area_t) return boolean is
  begin
    return area.firstCell.isNull;
  end;
  
  function makeRef (cell in cell_t) return varchar2 is
    str  varchar2(12);
    function makeRowRef return varchar2 is
    begin
      return case when cell.rw.isAbsolute then '$' end || to_char(cell.rw.value);
    end;
    function makeColRef return varchar2 is
    begin
      return case when cell.col.isAbsolute then '$' end || cell.col.alphaValue;
    end;
  begin
    if not cell.isNull then
      case cell.type
      when RT_COLUMN then
        str := makeColRef();
      when RT_ROW then
        str := makeRowRef();
      else
        str := makeColRef() || makeRowRef();
      end case;
    end if;
    return str;
  end;

  function makeArea (area in area_t) return varchar2
  is
  begin
    if not isNull(area) then
      return makeRef(area.firstCell) || ':' || makeRef(area.lastCell);
    else
      return '#NULL!';
    end if;
  end;
  
  function makeSheetPrefix (prefix in worksheetPrefix_t) return varchar2
  is
    str  varchar2(32767);
  begin
    if not prefix.isNull then
      str := prefix.sheetRange.firstSheet.name;
      if not prefix.sheetRange.isSingleSheet then
        str := str || ':' || prefix.sheetRange.lastSheet.name;
      end if;
      if prefix.sheetRange.firstSheet.quotable or prefix.sheetRange.lastSheet.quotable then
        str := '''' || replace(str, '''', '''''') || '''';
      end if;
      str := str || '!';
    end if;
    return str;
  end;
  
  function makeRef3d (ref3d in area3d_t) return varchar2
  is
  begin
    return makeSheetPrefix(ref3d.prefix) || makeRef(ref3d.area.firstCell);
  end;

  function makeArea3d (area3d in area3d_t) return varchar2
  is
  begin
    return makeSheetPrefix(area3d.prefix) || makeArea(area3d.area);
  end;

  function makeName (name in name_t) return varchar2
  is
    prefix  varchar2(256);
  begin
    if name.scope.idx is not null then
      prefix := name.scope.name;
      if name.scope.quotable then
        prefix := '''' || prefix || '''';
      end if;
      prefix := prefix || '!';
    end if;
    return prefix || name.value;
  end;
  
  function makeArray (array in array_t) return varchar2
  is
    str  varchar2(32767);
  begin
    str := '{';
    for r in 1 .. array.count loop
      if r > 1 then
        str := str || ';';
      end if;
      for i in 1 .. array(r).count loop
        if i > 1 then
          str := str || ',';
        end if;
        case 
        when array(r)(i).numValue is not null then
          str := str || to_char(array(r)(i).numValue, 'TM9', 'nls_numeric_characters=''.,''');
        when array(r)(i).strValue is not null then
          str := str || '"'||replace(array(r)(i).strValue, '"', '""')||'"';
        when array(r)(i).errValue is not null then
          str := str || array(r)(i).errValue;
        when array(r)(i).boolValue is not null then
          str := str || case when array(r)(i).boolValue = 1 then 'TRUE' else 'FALSE' end;
        else
          null;
        end case;
      end loop;
    end loop;
    str := str || '}';
    return str;
  end;
  
  procedure normalizePrefix (p in out nocopy worksheetPrefix_t)
  is
    tmp  sheet_t;
  begin
    if not p.isNull and not p.sheetRange.isSingleSheet then
      if p.sheetRange.firstSheet.idx > p.sheetRange.lastSheet.idx then
        -- swap sheet references
        tmp := p.sheetRange.firstSheet;
        p.sheetRange.firstSheet := p.sheetRange.lastSheet;
        p.sheetRange.lastSheet := tmp;
      end if;
    end if;
  end;
  
  procedure normalizeArea (a in out nocopy area_t)
  is
    tmp  cell_t;
  begin
    -- swap columns
    if a.firstCell.col.value > a.lastCell.col.value then
      tmp.col := a.firstCell.col;
      a.firstCell.col := a.lastCell.col;
      a.lastCell.col := tmp.col;
    end if;
    -- swap rows
    if a.firstCell.rw.value > a.lastCell.rw.value then
      tmp.rw := a.firstCell.rw;
      a.firstCell.rw := a.lastCell.rw;
      a.lastCell.rw := tmp.rw;
    end if;    
  end;

  procedure applyCellOffset (
    cell    in out nocopy cell_t
  , origin  in cell_t
  )
  is
  begin
    if not cell.col.isAbsolute then
      cell.col.value := cell.col.value - origin.col.value;
      if cell.col.value < 0 then
        cell.col.value := cell.col.value + MAX_COL_NUMBER;
      elsif cell.col.value >= MAX_COL_NUMBER then
        cell.col.value := cell.col.value - MAX_COL_NUMBER;
      end if;
      cell.col.alphaValue := base26Encode(cell.col.value + 1); 
    end if;
    if not cell.rw.isAbsolute then
      cell.rw.value := cell.rw.value - origin.rw.value;
      if cell.rw.value < 0 then
        cell.rw.value := cell.rw.value + MAX_ROW_NUMBER;
      elsif cell.rw.value >= MAX_ROW_NUMBER then
        cell.rw.value := cell.rw.value - MAX_ROW_NUMBER;
      end if;
      cell.rw.value := cell.rw.value + 1;
    end if;
  end;
  
  function createPtg (
    ptgType   in pls_integer
  , token     in varchar2 default null
  , pos       in pos_t default null
  ) 
  return pls_integer 
  is
    node  parseNode_t;
  begin
    t.extend;
    node.id := t.last;
    node.nodeType := ptgType;
    node.children := intList_t();
    node.token := token;
    node.pos := pos;
    --node.nodeClass := nullif(bitand(ptgType,96)/32, 0);
    node.nodeClass := bitand(ptgType,96)/32;
    t(node.id) := node;
    return node.id;
  end;

  function createPtgErr (errValue in varchar2, pos in pos_t) return pls_integer is
    nodeId  pls_integer := createPtg(PTG_ERR, errValue, pos);
  begin
    if errorMap.exists(errValue) then
      ptgErrList(nodeId).value := errorMap(errValue);
    else
      error('Invalid error value: %s', errValue);
    end if;
    return nodeId;
  end;

  function createPtgBool (boolValue in varchar2, pos in pos_t) return pls_integer is
    nodeId  pls_integer := createPtg(PTG_BOOL, boolValue, pos);
  begin
    case boolValue 
    when 'TRUE' then
      ptgBoolList(nodeId).value := 1;
    when 'FALSE' then
      ptgBoolList(nodeId).value := 0;
    else
      error('Invalid boolean value: %s', boolValue);
    end case;
    return nodeId;
  end;

  function createPtgName (nameValue in name_t, pos in pos_t) return pls_integer is
    localName  name_t := nameValue;
    nodeId     pls_integer;
    nameKey    varchar2(2048) := upper(ctx.sheet.name || '!' || localName.value);
  begin
    -- check if this name matches a defined name in the scope of the context sheet
    if ctx.nameMap.exists(nameKey) then
      localName := ctx.nameMap(nameKey);
    -- else should match a workbook-level name
    elsif ctx.nameMap.exists(upper(localName.value)) then
      localName := ctx.nameMap(upper(localName.value));
    else
      error('Undefined name: %s', localName.value);
    end if;
    nodeId := createPtg(PTG_NAME_R, makeName(localName), pos);
    ptgNameList(nodeId).value := localName;
    return nodeId;
  end;

  function createPtgNameX (nameValue in name_t, pos in pos_t) return pls_integer is
    nodeId     pls_integer;
    nameKey    varchar2(2048) := upper(nameValue.scope.name || '!' || nameValue.value);
    nameToken  varchar2(2048) := makeName(nameValue);
  begin
    if not ctx.nameMap.exists(nameKey) then
      error('Undefined name: %s', nameToken);
    end if;
    nodeId := createPtg(PTG_NAMEX_R, nameToken, pos);
    ptgNameList(nodeId).value := ctx.nameMap(nameKey);
    return nodeId;
  end;

  function createPtgRef (refValue in area3d_t, pos in pos_t) return pls_integer is
    nodeId  pls_integer := createPtg(PTG_REF_R, makeRef(refValue.area.firstCell), pos);
  begin
    ptgArea3dList(nodeId).value := refValue;
    return nodeId;
  end;

  function createPtgArea (areaValue in area3d_t, pos in pos_t) return pls_integer is
    nodeId  pls_integer := createPtg(PTG_AREA_R, makeArea(areaValue.area), pos);
  begin
    ptgArea3dList(nodeId).value := areaValue;
    return nodeId;
  end;

  function createPtgRef3d (ref3dValue in area3d_t, pos in pos_t) return pls_integer is
    nodeId      pls_integer;
    localRef3d  area3d_t := ref3dValue;
  begin
    if ctx.config.fmlaType = EXC_FMLATYPE_NAME then
      applyCellOffset(localRef3d.area.firstCell, ctx.cell);
    end if;
    nodeId := createPtg(PTG_REF3D_R, makeRef3d(localRef3d), pos);
    ptgArea3dList(nodeId).value := localRef3d;
    return nodeId;
  end;
  
  function createPtgArea3d (area3dValue in area3d_t, pos in pos_t) return pls_integer is
    nodeId       pls_integer;
    localArea3d  area3d_t := area3dValue;
  begin
    if ctx.config.fmlaType = EXC_FMLATYPE_NAME then
      applyCellOffset(localArea3d.area.firstCell, ctx.cell);
      applyCellOffset(localArea3d.area.lastCell, ctx.cell);
    end if;
    nodeId := createPtg(PTG_AREA3D_R, makeArea3d(localArea3d), pos);
    ptgArea3dList(nodeId).value := localArea3d;
    return nodeId;
  end;

  function createPtgStr (strValue in varchar2, pos in pos_t) return pls_integer is
    nodeId  pls_integer;
  begin
    if length(strValue) > 255 then
      error('String literals in formulas can''t be bigger than 255 characters');
    end if;
    nodeId := createPtg(PTG_STR, '"'||replace(strValue, '"', '""')||'"', pos);
    ptgStrList(nodeId).value := strValue;
    return nodeId;
  end;
    
  function createNumericPtg (strValue in varchar2, pos in pos_t) return pls_integer is
    nodeId    pls_integer;
    numValue  number := toLocalNumber(strValue);
  begin
    if trunc(numValue) = numValue then
      nodeId := createPtg(PTG_INT, to_char(numValue), pos);
      ptgIntList(nodeId).value := numValue;
    else
      nodeId := createPtg(PTG_NUM, replace(to_char(numValue), decimalSep, '.'), pos);
      ptgNumList(nodeId).value := to_binary_double(numValue);
    end if;
    return nodeId;
  end;

  function createPtgArray (arrayValue in array_t, pos in pos_t) return pls_integer is
    nodeId  pls_integer := createPtg(PTG_ARRAY_A, makeArray(arrayValue), pos);
  begin
    ptgArrayList(nodeId).value := arrayValue;
    return nodeId;
  end;

  function createPtgFuncVar (funcName in varchar2, pos in pos_t) return pls_integer is
    nodeId    pls_integer;
    nodeType  pls_integer;
    meta      functionMetadata_t;
  begin
    if not functionMetadataMap.exists(funcName) then
      error('Unknown or unsupported function call: ''%s''', funcName);
    end if;
    
    meta := functionMetadataMap(funcName);
    
    ctx.volatile := meta.volatile;
    
    nodeType := case when meta.minParamCount != meta.maxParamCount then PTG_FUNCVAR_R else PTG_FUNC_R end;
    applyPtgDataType(nodeType, meta.returnClass);
    
    nodeId := createPtg(nodeType, meta.internalName, pos);
    ptgFuncVarList(nodeId).name := funcName;
    return nodeId;
  end;

  function createPtgParen return parseNode_t is
    nodeId  pls_integer := createPtg(PTG_PAREN, '()');
  begin
    return t(nodeId);
  end;

  function createPtgMemErr (errValue in varchar2) return pls_integer is
    nodeId  pls_integer := createPtg(PTG_MEMERR_R);
  begin
    ptgMemErrList(nodeId).value := errorMap(errValue);
    return nodeId;
  end;

  function createPtgMemArea (areaListValue in areaList_t) return pls_integer is
    nodeId  pls_integer := createPtg(PTG_MEMAREA_R);
  begin
    ptgMemAreaList(nodeId).value := areaListValue;
    return nodeId;
  end;

  function parseArray (stream in out nocopy tokenStream_t) return array_t is
    arr       array_t := array_t();
    colCount  pls_integer;
    token     token_t;

    function accept (tokenType in pls_integer) return boolean is
    begin
      if token.type = tokenType then
        token := getNextToken(stream);
        return true;
      else
        return false;
      end if;
    end;
      
    procedure expect (tokenType in pls_integer) is
    begin
      if not accept(tokenType) then
        error('Unexpected token: %s', token.value);
      end if;
    end;

    procedure skipWhitespace is
    begin
      if token.type = T_WSPACE then
        token := getNextToken(stream);
      end if;
    end;

    function parseArrayItem return arrayItem_t is 
      item  arrayItem_t;
    begin
      skipWhitespace;
      case token.type
      when T_STRING then
        item.itemType := 1;
        item.strValue := token.value;
      when T_ERROR then
        item.itemType := 4;
        item.errValue := errorMap(token.value);
      when T_BOOLEAN then
        item.itemType := 2;
        item.boolValue := case when token.value = 'TRUE' then 1 else 0 end;
      when OP_MINUS then
        token := getNextToken(stream);
        if token.type = T_NUMBER then
          item.itemType := 0;
          item.numValue := toLocalNumber(token.value) * -1;
        else
          error('Unexpected token: %s', token.value);
        end if;
      when T_NUMBER then
        item.itemType := 0;
        item.numValue := toLocalNumber(token.value);
      else
        error('Unexpected token: %s', token.value);
      end case;
      return item;
    end;

    function parseArrayRow return arrayRow_t is
      rw  arrayRow_t := arrayRow_t();
    begin
      expect(T_ARRAYROW_START);
      loop
        rw.extend;
        rw(rw.last) := parseArrayItem();
        token := getNextToken(stream);
        skipWhitespace;
        exit when token.type = T_ARRAYROW_STOP;
        if token.type = T_ARRAYITEM_SEP then
          token := getNextToken(stream);
          continue;
        else
          error('Unexpected token: %s', token.value);
        end if;
      end loop;
      return rw;
    end;

  begin
    token := stream.tokens(stream.idx);
    expect(T_ARRAY_START);
    skipWhitespace;
    
    if token.type = T_ARRAYROW_STOP then
      error('Empty array');
    end if;
    
    loop
      arr.extend;
      arr(arr.last) := parseArrayRow();
      token := getNextToken(stream);
      exit when token.type = T_ARRAY_STOP;
      if token.type != T_ARRAYROW_START then
        error('Unexpected token: %s', token.value);
      end if;
    end loop;
    
    -- check if all rows have the same number of columns
    if arr.count != 0 then
      colCount := arr(1).count;
      for i in 2 .. arr.count loop
        if arr(i).count != colCount then
          error('Array row %d length mismatch: found %d, expected %d', i, arr(i).count, colCount);
        end if;
      end loop;
    end if;
    return arr;
  end;
  
  function base26Decode (input in varchar2) return pls_integer 
  result_cache
  is
    output  pls_integer;
    base    simple_integer := 1;
    c       varchar2(1 char);
  begin
    if input is not null then
      output := 0;
      for i in 1 .. length(input) loop
        c := substr(input, -i, 1);
        if not isLetter(c) then
          error('Invalid column reference: ''%s''', input);
        end if;
        output := output + (ascii(c) - 64) * base;
        base := base * 26;
      end loop;
    end if;
    return output;
  end;

  function base26Encode (input in pls_integer) 
  return varchar2
  result_cache
  is
    output  varchar2(3);
    num     pls_integer := input;
  begin
    if num is not null then
      while num != 0 loop
        output := chr(65 + mod(num-1,26)) || output;
        num := trunc((num-1)/26);
      end loop;
    end if;
    return output;
  end;
  
  function isExpr (node in parseNode_t) return boolean is
  begin
    return ( node.nodeType is null or node.nodeType > 255 );
  end;

  function isPtgFunc (ptgType in pls_integer) return boolean is
  begin
    return ( ptgType in (PTG_FUNC_R, PTG_FUNCVAR_R, PTG_FUNC_V, PTG_FUNCVAR_V, PTG_FUNC_A, PTG_FUNCVAR_A) );
  end;
  
  function isAuxPtg (ptgType in pls_integer) return boolean is
  begin
    return ( ptgType in (PTG_PAREN, PTG_MEMFUNC_A, PTG_MEMFUNC_R, PTG_MEMFUNC_V
                                  , PTG_MEMAREA_A, PTG_MEMAREA_R, PTG_MEMAREA_V
                                  , PTG_MEMERR_A, PTG_MEMERR_R, PTG_MEMERR_V) );
  end;
  
  function isRef (ptg in parseNode_t, matchFunctionEnd in pls_integer) return boolean is
  begin
    return ptg.nodeType in (PTG_AREA_R, PTG_AREA3D_R, PTG_REF_R, PTG_REF3D_R, PTG_NAME_R, PTG_NAMEX_R)
        -- convoluted stuff to test whether the node matches a ref function start or stop, depending on the calling context
        or (
             (
               matchFunctionEnd = MATCH_FUNC_START and isPtgFunc(ptg.nodeType)
               or 
               matchFunctionEnd = MATCH_FUNC_STOP and ptg.nodeType = XPTG_FUNC_STOP
             ) 
             and ptg.token in ('OFFSET','INDEX','CHOOSE','IF','INDIRECT')
           );
  end;
  
  function isBasePtg (ptgType in pls_integer) return boolean is
  begin
    -- bits 5 and 6 hold the Ptg data type if applicable : 
    -- 1 = REFERENCE, 2 = VALUE, 3 = ARRAY
    return ( bitand(ptgType,96)/32 = 0 );
  end;
  
  function parseSheetPrefix (input in varchar2) return worksheetPrefix_t is
    str     varchar2(32767);
    prefix  worksheetPrefix_t;
  begin
    if instr(input, ':') != 0 then
        
      -- sheet range
      prefix.sheetRange.isSingleSheet := false;
      -- validate first sheet
      str := substr(input, 1, instr(input, ':') - 1);
      if ctx.sheetMap.exists(upper(str)) then
        prefix.sheetRange.firstSheet := ctx.sheetMap(upper(str));
      else
        error('Undefined sheet name: %s', str);
      end if;
      -- validate last sheet
      str := substr(input, instr(input, ':') + 1);
      if ctx.sheetMap.exists(upper(str)) then
        prefix.sheetRange.lastSheet := ctx.sheetMap(upper(str));         
      else
        error('Undefined sheet name: %s', str);
      end if;
        
    else

      -- single sheet
      if ctx.sheetMap.exists(upper(input)) then
        prefix.sheetRange.firstSheet := ctx.sheetMap(upper(input));
      else
        error('Undefined sheet name: %s', input);
      end if;        
        
    end if;
    prefix.isNull := false;
    return prefix;
  end;

  function parseCellReference (input in varchar2) return cell_t is
    cell  cell_t;

    absolute  boolean := false;
    str       varchar2(32767);
    i         pls_integer := 0;
    c         varchar2(1 char);
    procedure get is 
    begin
      i := i + 1;
      c := substr(input, i, 1);
    end;

  begin
    
    get;
    
    if c = '$' then
      absolute := true;
      get;
    end if;
    
    if isLetter(c) then
      str := str || c;
      get;
      while isLetter(c) loop
        str := str || c;
        get;
      end loop;
      
      if length(str) > 3 then
        return NULL_CELL;
      end if;
      
      cell.col.value := base26Decode(str);
      if cell.col.value between 1 and MAX_COL_NUMBER then
        cell.col.isAbsolute := absolute;
        cell.col.alphaValue := str;
        cell.type := RT_COLUMN;
        cell.isNull := false;
      else
        return NULL_CELL;
      end if;
      
      if c is null then  
        return cell;
      end if;
      
      absolute := false;
      str := null;
      
      if c = '$' then
        absolute := true;
        get;
      end if;
      
    end if;
        
    if isDigit(c) and c != '0' then
      str := str || c;
      get;
      while isDigit(c) loop
        str := str || c;
        get;
      end loop;
      -- extra characters found after sequence of digits
      if c is not null then
        return NULL_CELL;
      end if;
      cell.rw.value := toLocalNumber(str);
      if cell.rw.value between 1 and MAX_ROW_NUMBER then
        cell.rw.isAbsolute := absolute;
        cell.type := cell.type + RT_ROW;
        cell.isNull := false;
        return cell;
      else
        return NULL_CELL;
      end if;
    end if;
    
    return NULL_CELL;
    
  end;

  function isValidCellReference (input in varchar2) return boolean is
    cell  cell_t := parseCellReference(input);
  begin
    return not cell.isNull and cell.type = RT_CELL;
  end;
  
  function tokenize (input in varchar2) return tokenStream_t
  is
  
    inString    boolean := false;
    inPath      boolean := false;
    inError     boolean := false;
    token       token_t;
    stream      tokenStream_t;
    stack       tokenStream_t;
    fnToken     token_t;
    startPos    pos_t;
    
    procedure skipWs is
    begin
      --getChar;
      while look = ' ' and not eof() loop
        getChar;
      end loop;    
    end;
    
    procedure trimWs is
    begin
      if peekToken(stream).type = T_WSPACE then
        stream.tokens.trim;
        stream.idx := stream.idx - 1;
      end if;
    end;
  
  begin
    
    --normalize input
    formulaString := replace(input, chr(13)||chr(10), chr(10));
    formulaString := replace(formulaString, chr(13), chr(10));
    formulaString := trim(formulaString);
    
    formulaLength := length(formulaString);
    pointer := 0;
    pos.ln := 1;
    pos.cn := 0;
    
    getChar;
    
    while not eof() loop

      if look = chr(10) then
        getChar;
        continue;
      end if;
    
      if inString then
        if look = '"' then
          if nextChar() = '"' then
            -- escape character
            append(token, look);
            getChar;
          else
            -- end of string reached
            inString := false;
            pushToken(stream, token);
            token := null;
          end if;
        else
          append(token, look); 
        end if;
        getChar;
        continue;
      end if;
      
      if inPath then
        if look = '''' then
          if nextChar() = '''' then
            -- escape character
            append(token, look);
            getChar;
          else
            -- end of path reached
            inPath := false;
            -- trim leading apostrophe
            token.value := substr(token.value, 2);
            -- try to parse a worksheet prefix
            token.parsedValue.prefix := parseSheetPrefix(token.value);
            token.type := T_PREFIX;
            
            pushToken(stream, token);
            token := null;
          end if;
        else
          append(token, look);
        end if;
        getChar;
        continue;
      end if;
      
      if inError then
        append(token, look);
        getChar;
        if token.value in ('#NULL!','#DIV/0!','#VALUE!','#REF!','#NAME?','#NUM!','#N/A') then
          inError := false;
          pushToken(stream, token);
          token := null;
        end if;
        continue;
      end if;

      if look = ',' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        case peekToken(stack).type
        when T_FUNC_START then
          -- insert a missing-argument token
          if previousNonWsToken(stream).type in (T_FUNC_START, T_ARG_SEP) then
            pushToken(stream, null, T_MISSING_ARG, pos);
          end if;
          pushToken(stream, ',', T_ARG_SEP, pos);
        when T_ARRAYROW_START then
          pushToken(stream, ',', T_ARRAYITEM_SEP, pos);
        else
          pushToken(stream, ',', OP_UNION, pos);
        end case;
        getChar;
        skipWs;
        continue;
      end if;      
    
      if isDigit(look) and ( token.value is null /*or isPrevNonDigitBlank()*/ ) and not isPrevOrNextNonDigitRangeOp() then
        
        token.type := T_NUMBER;
        token.pos := pos;
        
        append(token, look);
        getChar;
        while isDigit(look) loop
          append(token, look);
          getChar;
        end loop;
        
        -- decimal separator
        if look = '.' then
          append(token, look);
          getChar;
          -- decimal part
          while isDigit(look) loop
            append(token, look);
            getChar;
          end loop;
        end if;
        
        -- scientific notation
        if look in ('E','e') then
          append(token, look);
          getChar;
          -- exponent sign
          if look in ('+','-') then
            append(token, look);
            getChar;
          end if;
          -- exponent
          while isDigit(look) loop
            append(token, look);
            getChar;
          end loop;
        end if;
        
        pushToken(stream, token);
        token := null;
        
        continue;
        
      end if;

      if look = '"' then
        trimWs;
        if token.value is not null then
          error('Unexpected token: %s', token.value);
        end if;
        token := newToken(T_STRING, tokenPos => pos);
        inString := true;
        getChar;
        continue;
      end if;

      if look = '''' then
        if token.value is not null then
          error('Unexpected token: %s', token.value);
        end if;
        token := newToken(T_QUOTED, look, pos);
        inPath := true;
        getChar;
        continue;
      end if;      
      
      if look = '#' then
        trimWs;
        if token.value is not null then
          error('Unexpected token: %s', token.value);
        end if;
        inError := true;
        token := newToken(T_ERROR, look, pos);
        getChar;
        continue;
      end if;

      if look = '{' then
        trimWs;
        if token.value is not null then
          error('Unexpected token: %s', token.value);
        end if;

        token := null;
        
        pushToken(stack, pushToken(stream, look, T_ARRAY_START, pos));
        pushToken(stack, pushToken(stream, null, T_ARRAYROW_START, pos));
        
        getChar; 
        skipWs;
        continue;
      end if;
      
      if look = ';' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        
        if peekToken(stack).type = T_ARRAYROW_START then
          popToken(stack);
          pushToken(stream, null, T_ARRAYROW_STOP, pos);
          pushToken(stack, pushToken(stream, null, T_ARRAYROW_START, pos));
        else
          error('Unexpected token: %s', look);
        end if;
        
        getChar;
        skipWs;
        continue;
      
      end if;
      
      if look = '}' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        
        if peekToken(stack).type = T_ARRAYROW_START then
          popToken(stack);
          pushToken(stream, null, T_ARRAYROW_STOP, pos);
          popToken(stack);
          pushToken(stream, null, T_ARRAY_STOP, pos);
        else
          error('Unexpected token: %s', look);
        end if;
        
        getChar;
        skipWs;
        continue;
      
      end if;

      if look = ':' then
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, OP_RANGE, pos);
        getChar;
        continue;
      end if;  

      if look = '!' then
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, T_BANG, pos);
        getChar;
        skipWs;
        continue;
      end if;
      
      if look = ' ' then
        
        pushAndClear(stream, token, T_OPERAND);
        
        pushToken(stream, look, T_WSPACE, pos);
        getChar;
        skipWs;
        continue;
      
      end if;

      if look = '>' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        startPos := pos;
        getChar;
        if look = '=' then
          pushToken(stream, '>=', OP_GE, startPos);
          getChar;
        else
          pushToken(stream, '>', OP_GT, startPos);
        end if;
        skipWs;
        continue;
      end if;
      
      if look = '<' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        startPos := pos;
        getChar;
        if look = '=' then
          pushToken(stream, '<=', OP_LE, startPos);
          getChar;
        elsif look = '>' then
          pushToken(stream, '<>', OP_NE, startPos);
          getChar;
        else
          pushToken(stream, '<', OP_LT, startPos);
        end if;
        skipWs;
        continue;
      end if;
      
      if look = '+' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, OP_PLUS, pos);
        getChar;
        skipWs;
        continue;
      end if;      

      if look = '-' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, OP_MINUS, pos);
        getChar;
        skipWs;
        continue;
      end if;

      if look = '*' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, OP_MUL, pos);
        getChar;
        skipWs;
        continue;
      end if;

      if look = '/' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, OP_DIV, pos);
        getChar;
        skipWs;
        continue;
      end if;

      if look = '^' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, OP_EXP, pos);
        getChar;
        skipWs;
        continue;
      end if;

      if look = '&' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, OP_CONCAT, pos);
        getChar;
        skipWs;
        continue;
      end if;

      if look = '=' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, OP_EQ, pos);
        getChar;
        skipWs;
        continue;
      end if;

      if look = '%' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        pushToken(stream, look, OP_PERCENT, pos);
        getChar;
        skipWs;
        continue;
      end if;
      
      if look = '(' then
        fnToken := pushAndClear(stream, token, T_FUNC_START);
        if fnToken.value is not null then         
          pushToken(stack, fnToken);
        else
          trimWs;
          pushToken(stack, pushToken(stream, null, T_LEFT, pos));
        end if;
        getChar;
        skipWs;
        continue;
      end if;
      
      if look = ')' then
        trimWs;
        pushAndClear(stream, token, T_OPERAND);
        case peekToken(stack).type
        when T_FUNC_START then
          -- insert a missing-argument token
          if previousNonWsToken(stream).type = T_ARG_SEP then
            pushToken(stream, null, T_MISSING_ARG, pos);
          end if;
          pushToken(stream, peekToken(stack).value, T_FUNC_STOP, pos);
          popToken(stack);
        when T_LEFT then
          popToken(stack);
          pushToken(stream, null, T_RIGHT, pos);
        else
          error('Unexpected token: %s', look);
        end case;
        getChar;
        skipWs;
        continue;
      end if;
      
      append(token, look);
      getChar;
    
    end loop;
    
    pushAndClear(stream, token, T_OPERAND);
    
    for i in 1 .. stream.tokens.count loop
      dbms_output.put_line(
        utl_lms.format_message(
          '[%d][%s](%d,%d) %s'
        , stream.tokens(i).type
        , tokenLabels(stream.tokens(i).type)
        , stream.tokens(i).pos.ln
        , stream.tokens(i).pos.cn
        , stream.tokens(i).value
        )
      );
    end loop;
    dbms_output.put_line('===============================================================');
    return stream;
    
  end;
  
  procedure extendArea (a in out nocopy area_t) is
  begin
    -- column range: add row info
    if a.firstCell.type = RT_COLUMN then
      a.firstCell.rw.value := 1;
      a.lastCell.rw.value := MAX_ROW_NUMBER;
    -- row range: add column info
    elsif a.firstCell.type = RT_ROW then
      a.firstCell.col.value := 1;
      a.lastCell.col.value := MAX_COL_NUMBER;
    -- single cell ref: copy first cell to last cell
    elsif a.lastCell.isNull then
      a.lastCell := a.firstCell;
    end if;    
  end;

  -- compute intersection of two areas
  function evalIsect (a in area_t, b in area_t) return area_t
  is
    r  area_t;
  begin
    if not(isNull(a) or isNull(b)) then
      
      if not ( 
           a.lastCell.col.value < b.firstCell.col.value 
        or a.firstCell.col.value > b.lastCell.col.value
        or a.lastCell.rw.value < b.firstCell.rw.value 
        or a.firstCell.rw.value > b.lastCell.rw.value 
        ) 
      then
        
        r.firstCell.col.value := greatest(a.firstCell.col.value, b.firstCell.col.value);
        r.firstCell.col.alphaValue := base26Encode(r.firstCell.col.value);
        r.firstCell.rw.value := greatest(a.firstCell.rw.value, b.firstCell.rw.value);
        r.firstCell.isNull := false;
        
        r.lastCell.col.value := least(a.lastCell.col.value, b.lastCell.col.value);
        r.lastCell.col.alphaValue := base26Encode(r.lastCell.col.value);
        r.lastCell.rw.value := least(a.lastCell.rw.value, b.lastCell.rw.value);
        r.lastCell.isNull := false;
      
      end if;
    end if;
    
    return r;
    
  end;
  
  -- compute the combined range of multiple areas
  function evalRange (areas in areaList_t) return area_t
  is
    r  area_t;
  begin
    
    if areas is not empty then
      
      r := areas(1);
      if not isNull(r) then
        for i in 2 .. areas.count loop
          if not isNull(areas(i)) then
            r.firstCell.col.value := least(r.firstCell.col.value, areas(i).firstCell.col.value);
            r.firstCell.col.alphaValue := base26Encode(r.firstCell.col.value);
            r.firstCell.rw.value := least(r.firstCell.rw.value, areas(i).firstCell.rw.value);
            r.firstCell.isNull := false;
              
            r.lastCell.col.value := greatest(r.lastCell.col.value, areas(i).lastCell.col.value);
            r.lastCell.col.alphaValue := base26Encode(r.lastCell.col.value);
            r.lastCell.rw.value := greatest(r.lastCell.rw.value, areas(i).lastCell.rw.value);
            r.lastCell.isNull := false;
          else
            r := NULL_AREA;
            exit;
          end if;
        end loop;
      end if;
      
    end if;
        
    return r;
    
  end;

  -- concatenate two area lists, preserving order
  function evalUnion (a in areaList_t, b in areaList_t) return areaList_t
  is
    areas  areaList_t := a;
  begin
    -- could have used MULTISET UNION operator here but not sure the order is preserved
    areas.extend(b.count);
    for i in 1 .. b.count loop
      areas(a.count + i) := b(i);
    end loop;
    return areas;
  end;
  
  procedure test_eval_isect (r1 in varchar2, r2 in varchar2) is
    a1  area_t;
    a2  area_t;
    
    procedure setArea (a in out nocopy area_t, input in varchar2) is
      p  pls_integer := instr(input, ':');
    begin
      if p != 0 then
        a.firstCell := parseCellReference(substr(input, 1, p-1));
        a.lastCell := parseCellReference(substr(input, p+1));
      else
        a.firstCell := parseCellReference(input);
      end if;      
    end;
    
  begin
    setArea(a1, r1);
    setArea(a2, r2);
    
    dbms_output.put_line(makeArea(evalIsect(a1,a2)));
    
  end;

  procedure test_eval_range (list in apex_t_varchar2) is
    areas  areaList_t := areaList_t();
    area   area_t;
    
    procedure setArea (a in out nocopy area_t, input in varchar2) is
      p  pls_integer := instr(input, ':');
    begin
      if p != 0 then
        a.firstCell := parseCellReference(substr(input, 1, p-1));
        a.lastCell := parseCellReference(substr(input, p+1));
      else
        a.firstCell := parseCellReference(input);
      end if;      
    end;
    
  begin
    
    if list is not empty then
      areas.extend(list.count);
      for i in 1 .. list.count loop
        area := NULL_AREA;
        setArea(area, list(i));
        extendArea(area);
        areas(i) := area;
      end loop;
    end if;
    
    dbms_output.put_line(makeArea(evalRange(areas)));
    
  end;

  function getRPNArray (rootId in pls_integer, includeAux in boolean default true) return intList_t
  is
    arr  intList_t := intList_t();
    procedure readNode (nodeId in pls_integer) is
    begin
      if isExpr(t(nodeId)) then
        for i in 1 .. t(nodeId).children.count loop
          readNode(t(nodeId).children(i));
        end loop;
      else
        if not isAuxPtg(t(nodeId).nodeType) or includeAux then
          arr.extend;
          arr(arr.last) := nodeId;
          dbms_output.put_line(ptgLabel(t(nodeId)));
        end if;
      end if;
    end;
  begin

    dbms_output.put_line('======= BEGIN RPN ARRAY =======');
    readNode(rootId);
    dbms_output.put_line('======= END RPN ARRAY =======');
    
    return arr;
  end;

  function evalBinaryRef (rootId in pls_integer) return areaList_t
  is
    type areaListStack_t is table of areaList_t;
    st  areaListStack_t := areaListStack_t();
    
    ptgArray  intList_t := getRPNArray(rootId, false);
    
    area   area_t;
    areas  areaList_t;
    ptg    parseNode_t;
    a1     areaList_t;
    a2     areaList_t;
    
    idx  pls_integer := 0;
    
  begin
    
    while idx < ptgArray.count loop
      idx := idx + 1;
      ptg := t(ptgArray(idx));
        
      case
      when ptg.nodeType in (PTG_REF_A, PTG_REF_R, PTG_REF_V, PTG_AREA_A, PTG_AREA_R, PTG_AREA_V) then
        area := ptgArea3dList(ptg.id).value.area;
        extendArea(area);
        st.extend;
        st(st.last) := areaList_t(area);
        
      when ptg.nodeType = PTG_UNION then
        -- concat top-2 lists and push back the result
        areas := evalUnion(st(st.last-1), st(st.last));
        st.trim;
        st(st.last) := areas;
        
      when ptg.nodeType = PTG_ISECT then
        areas := areaList_t();
        a1 := st(st.last-1);
        a2 := st(st.last);
        for i in 1 .. a1.count loop
          for j in 1 .. a2.count loop
            area := evalIsect(a1(i), a2(j));
            --if not isNull(area) then
              areas.extend;
              areas(areas.last) := area;
            --end if;
          end loop;
        end loop;
        
        st.trim;
        st(st.last) := areas;
        
      when ptg.nodeType = PTG_RANGE then
        area := evalRange(st(st.last-1) multiset union st(st.last));
        st.trim;
        st(st.last) := areaList_t(area);
        
      else
        -- Anything else than a PtgRef, PtgArea, PtgUnion, PtgIsect or PtgRange means that
        -- the expression cannot be evaluated statically.
        -- Return an empty list
        st.delete;
        exit;
        
      end case; 
    
    end loop;
    
    if st is not empty then
      areas := st(1);
    else
      areas := areaList_t();
    end if;
    
    for i in 1 .. areas.count loop
      dbms_output.put_line(makeArea(areas(i)));
    end loop;
    
    return areas;
    
  end;

  procedure traverse (nodeId in out nocopy pls_integer, childIdx in pls_integer default null)
  is
    exprId    pls_integer;
    memPtgId  pls_integer;
    parentId  pls_integer;
    childId   pls_integer;
    areas     areaList_t;
  begin
  
    if t(nodeId).nodeType = EXPR_BINARY_REF then
      
      -- evaluate this expression
      areas := evalBinaryRef(nodeId);
      
      if areas is empty then
        -- expression cannot be evaluated statically, use PtgMemFunc
        memPtgId := createPtg(PTG_MEMFUNC_R);
      elsif isNull(areas(1)) then
        -- expression evaluates to #NULL!
        memPtgId := createPtgMemErr('#NULL!');
      else
        memPtgId := createPtgMemArea(areas);
      end if;
      
      exprId := createPtg(EXPR_MEM_AREA);
      t(exprId).nodeClass := VT_REFERENCE;
      t(exprId).operandInfo := t(nodeId).operandInfo;
      
      parentId := t(nodeId).parentId;
      
      -- attach new expr
      t(exprId).parentId := parentId;
      if parentId is not null then
        t(parentId).children(childIdx) := exprId;
      end if;
          
      t(exprId).children.extend(2);
          
      -- attach new mem Ptg
      t(exprId).children(1) := memPtgId;
      t(memPtgId).parentId := exprId;
          
      -- attach binary ref expr to new mem area expr
      t(nodeId).parentId := exprId;
      t(exprId).children(2) := nodeId;
      
      -- make this expression the new root of the current branch
      nodeId := exprId;
      
    else
      
      if t(nodeId).children is not empty then
        for i in 1 .. t(nodeId).children.count loop
          childId := t(nodeId).children(i);
          traverse(childId, i);
          t(nodeId).children(i) := childId;
        end loop;
      end if;
      
    end if;
    
  end;

  procedure transformOperand (
    nodeId         in pls_integer
  , convInfo       in operandInfo_t
  , prevConv       in pls_integer
  , prevClassConv  in pls_integer
  , wasRefClass    in boolean
  )
  is
  
    node        parseNode_t := t(nodeId);
    tokenClass  pls_integer := node.nodeClass;
    conv        pls_integer;
    classConv   pls_integer;
    childNodeId  pls_integer;
    nodeType     pls_integer;
  
  begin
    
    dbms_output.put_line(ptgLabel(node)||':'||node.token);
    
    if not isBasePtg(node.nodeType) or isExpr(node) then
    
      -- REF tokens in VALTYPE parameters behave like VAL tokens
      if convInfo.valueType and tokenClass = VT_REFERENCE then
        tokenClass := VT_VALUE;
        --apply
      end if;
      
      -- replace RPO conversion of operator with parent conversion
      conv := case when convInfo.convClass = EXC_PARAMCONV_RPO then prevConv else convInfo.convClass end;
      
      -- find the effective token class conversion to be performed for this token
      classConv := EXC_CLASSCONV_ORG;
      
      case conv
      when EXC_PARAMCONV_ORG then
        -- conversion is forced independent of parent conversion
        classConv := EXC_CLASSCONV_ORG;
        
      when EXC_PARAMCONV_VAL then
        -- conversion is forced independent of parent conversion
        classConv := EXC_CLASSCONV_VAL;
        
      when EXC_PARAMCONV_ARR then
        -- conversion is forced independent of parent conversion
        classConv := EXC_CLASSCONV_ARR;
        
      when EXC_PARAMCONV_RPT then
        case 
        when prevConv in (EXC_PARAMCONV_ORG, EXC_PARAMCONV_VAL, EXC_PARAMCONV_ARR) then
          /*  If parent token has REF class (REF token in REFTYPE
                          function parameter), then RPT does not repeat the
                          previous explicit ORG or ARR conversion, but always
                          falls back to VAL conversion. */
          classConv := case when wasRefClass then EXC_CLASSCONV_VAL else prevClassConv end;
          
        when prevConv = EXC_PARAMCONV_RPT then
          -- nested RPT repeats the previous effective conversion
          classConv := prevClassConv;
          
        when prevConv = EXC_PARAMCONV_RPX then
          /*  If parent token has REF class (REF token in REFTYPE
                          function parameter), then RPX repeats the previous
                          effective conversion (which will be either ORG or ARR,
                          but never VAL), otherwise falls back to ORG conversion. */
          classConv := case when wasRefClass then prevClassConv else EXC_CLASSCONV_ORG end;
        
        when prevConv = EXC_PARAMCONV_RPO then
          null;
          
        end case;
        
      when EXC_PARAMCONV_RPX then
        /*  If current token still has REF class, set previous effective
                  conversion as current conversion. This will not have an effect
                  on the REF token but is needed for RPT parameters of this
                  function that want to repeat this conversion type. If current
                  token is VAL or ARR class, the previous ARR conversion will be
                  repeated on the token, but VAL conversion will not. */
        classConv := case when tokenClass = VT_REFERENCE or prevClassConv = EXC_CLASSCONV_ARR then prevClassConv else EXC_CLASSCONV_ORG end;
      
      when EXC_PARAMCONV_RPO then
        null;
        
      end case;
      
      -- do the token class conversion                
      case classConv
      when EXC_CLASSCONV_ORG then
        /*  Cell formulas: leave the current token class. Cell formulas
                  are the only type of formulas where all tokens can keep
                  their original token class.
                  Array and defined name formulas: convert VAL to ARR. */
        if ctx.config.fmlaClassType != EXC_CLASSTYPE_CELL and tokenClass = VT_VALUE then
          tokenClass := VT_ARRAY;
          --apply
        end if;
        
      when EXC_CLASSCONV_VAL then
        if tokenClass = VT_ARRAY then
          tokenClass := VT_VALUE;
          --apply
        end if;
        
      when EXC_CLASSCONV_ARR then
        if tokenClass = VT_VALUE then
          tokenClass := VT_ARRAY;
          -- apply
        end if;
        
      end case;
      
      
      if node.nodeType = EXPR_FUNC_CALL then
        t(nodeId).nodeClass := tokenClass;
        -- set class on the underlying Ptg
        childNodeId := node.children(node.children.last);
        nodeType := t(childNodeId).nodeType;
        -- if not PtgAttrSum
        if nodeType != PTG_ATTR_B then
          applyPtgDataType(nodeType, tokenClass);
          t(childNodeId).nodeType := nodeType;
          t(childNodeId).nodeClass := tokenClass;
        end if;
      
      elsif node.nodeType in (EXPR_BINARY_VAL, EXPR_BINARY_REF, EXPR_UNARY, EXPR_DISPLAY_PREC) then
        
        t(nodeId).nodeClass := tokenClass;
        
      elsif node.nodeType = EXPR_MEM_AREA then
        
        t(nodeId).nodeClass := tokenClass;
        -- apply token class to underlying mem Ptg (i.e. 1st child)
        childNodeId := node.children(1);
        nodeType := t(childNodeId).nodeType;
        applyPtgDataType(nodeType, tokenClass);
        t(childNodeId).nodeType := nodeType;
        t(childNodeId).nodeClass := tokenClass;
        
      else
        -- non-base Ptg
        applyPtgDataType(node.nodeType, tokenClass);
        t(nodeId).nodeType := node.nodeType;
        t(nodeId).nodeClass := tokenClass;
      
      end if;
      
      -- do not recurse on children of a Mem Area Expression as they retain their original token class
      if node.nodeType != EXPR_MEM_AREA then
        -- children
        for i in 1 .. node.children.count - 1 loop
          
          transformOperand(
            nodeId         => node.children(i)
          , convInfo       => case when t(node.children(i)).operandInfo.convClass is null then convInfo else t(node.children(i)).operandInfo end
          , prevConv       => conv
          , prevClassConv  => classConv
          , wasRefClass    => ( tokenClass = VT_REFERENCE )
          );
        
        end loop;
      end if;

    
    end if;
    
  
  end;

  procedure transformOperands (rootId in pls_integer) is
  
    nameFmla        boolean;
    paramConv       pls_integer;
    classConv       pls_integer;
    convInfo        operandInfo_t;
  
  begin
    
    if ctx.config.fmlaType is null then
      ctx.config.fmlaType := EXC_FMLATYPE_CELL;
    end if;
    ctx.config.fmlaClassType := case 
                                when ctx.config.fmlaType in (EXC_FMLATYPE_CELL, EXC_FMLATYPE_SHARED) then EXC_CLASSTYPE_CELL
                                when ctx.config.fmlaType in (EXC_FMLATYPE_MATRIX, EXC_FMLATYPE_CONDFMT, EXC_FMLATYPE_DATAVAL) then EXC_CLASSTYPE_ARRAY
                                else EXC_CLASSTYPE_NAME
                                end;
    nameFmla := ( ctx.config.fmlaClassType = EXC_CLASSTYPE_NAME );
    paramConv := case when nameFmla then EXC_PARAMCONV_ARR else EXC_PARAMCONV_VAL end;
    classConv := case when nameFmla then EXC_CLASSCONV_ARR else EXC_CLASSCONV_VAL end;    
  
    convInfo.convClass := paramConv;
    convInfo.valueType := not(nameFmla);  
  
    transformOperand(
      nodeId         => rootId
    , convInfo       => convInfo
    , prevConv       => paramConv
    , prevClassConv  => classConv
    , wasRefClass    => nameFmla
    );    
  
  end;

  function compile (input in intList_t) return pls_integer is

    rootId           pls_integer;
    currentParentId  pls_integer;
    firstChildId     pls_integer;
    
    node             parseNode_t;
    top              parseNode_t; 
    st               parseNodeList_t := parseNodeList_t();
    
    procedure setChildrenConversion (parentId in pls_integer, nodeType in pls_integer, funcName in varchar2 default null) is
      info   operandInfo_t;
      meta   functionMetadata_t;
      argc   pls_integer;
      
      procedure setRecurse (nodeId in pls_integer) is
      begin
        t(nodeId).operandInfo := info;
        if t(nodeId).nodeType = EXPR_DISPLAY_PREC then
          setRecurse(t(nodeId).children(1));
        end if;
      end;
      
    begin
      -- assign token class conversion type to operands i.e. all left siblings of this operator node
      if isPtgFunc(nodeType) then
        
        meta := functionMetadataMap(funcName);
        argc := t(parentId).children.count - 1;
        
        if argc < meta.minParamCount then
          error('Invalid number of parameters for function %s: expected at least %d, found %d', funcName, meta.minParamCount, argc);
        end if;
        
        if meta.maxParamCount != -1 and argc > meta.maxParamCount then
          error('Invalid number of parameters for function %s: expected at most %d, found %d', funcName, meta.maxParamCount, argc);
        end if;
      
        for i in 1 .. argc loop
          -- reuse the last param info if actual argc exceeds the number of defined parameters in function meta
          if i > meta.paramInfos.count then
            null;
          else
            info := meta.paramInfos(i);
          end if;
          setRecurse(t(parentId).children(i));
        end loop;
      
      else
        
        info := opMetadataMap(nodeType).operandInfo;
        for i in 1 .. t(parentId).children.count - 1 loop
          setRecurse(t(parentId).children(i));
        end loop;
        
      end if;
      
    end;
    
    -- push a node to the output tree
    procedure pushOut (node in parseNode_t) is
      childIdx     pls_integer;
      leftSibling  parseNode_t;
    begin

      childIdx := t(currentParentId).children.last;

      -- check if left sibling is a single-operand or empty expression, but not a function call
      if childIdx is not null then
        leftSibling := t(t(currentParentId).children(childIdx));
        if isExpr(leftSibling) and leftSibling.children.count > 1 and leftSibling.nodeType is null then
          -- anonymous expression with more than one child = syntax error
          error('Unexpected token: %s', t(leftSibling.children(2)).token);
        elsif isExpr(leftSibling) and leftSibling.children.count <= 1 and nvl(leftSibling.nodeType, -1) != EXPR_FUNC_CALL then
          if leftSibling.children.count = 1 then
            -- move operand
            t(leftSibling.children(1)).parentId := currentParentId;
            t(currentParentId).children(childIdx) := leftSibling.children(1);
          else
            -- remove empty expression from child list
            t(currentParentId).children.trim;
          end if;
          -- drop orphan expression node
          t.delete(leftSibling.id);
        end if;
      end if;

      t(node.id).parentId := currentParentId;     
      t(currentParentId).children.extend;
      childIdx := t(currentParentId).children.last;
      t(currentParentId).children(childIdx) := node.id;
      
      -- assign expression type to current parent when pushing an operator etc.
      if node.nodeType in (PTG_ADD, PTG_SUB, PTG_MUL, PTG_DIV, PTG_POWER, PTG_CONCAT, PTG_LT, PTG_LE, PTG_EQ, PTG_GE, PTG_GT, PTG_NE) then
        t(currentParentId).nodeType := EXPR_BINARY_VAL;
        t(currentParentId).nodeClass := opMetadataMap(node.nodeType).returnClass;
        setChildrenConversion(currentParentId, node.nodeType);
        
      elsif node.nodeType in (PTG_ISECT, PTG_UNION, PTG_RANGE) then
        
        if not(t(t(currentParentId).children(1)).nodeClass = VT_REFERENCE and t(t(currentParentId).children(2)).nodeClass = VT_REFERENCE) then
          error('Unexpected token: %s', ptgLabel(node));
        end if;
      
        t(currentParentId).nodeType := EXPR_BINARY_REF;
        t(currentParentId).nodeClass := opMetadataMap(node.nodeType).returnClass;
        setChildrenConversion(currentParentId, node.nodeType);
        
      elsif node.nodeType in (PTG_UPLUS, PTG_UMINUS, PTG_PERCENT) then
        t(currentParentId).nodeType := EXPR_UNARY;
        t(currentParentId).nodeClass := opMetadataMap(node.nodeType).returnClass;
        setChildrenConversion(currentParentId, node.nodeType);
        
      elsif isPtgFunc(node.nodeType) then
        t(currentParentId).nodeType := EXPR_FUNC_CALL;
        t(currentParentId).nodeClass := node.nodeClass;
        setChildrenConversion(currentParentId, node.nodeType, ptgFuncVarList(node.id).name);
        
      elsif node.nodeType = PTG_PAREN then
        t(currentParentId).nodeType := EXPR_DISPLAY_PREC;
        -- apply node class of the first child (i.e. the encapsulated expression), VALUE by default
        --t(currentParentId).nodeClass := nvl(t(t(currentParentId).children(1)).nodeClass, VT_VALUE);
        t(currentParentId).nodeClass := case when t(t(currentParentId).children(1)).nodeClass = VT_NONE 
                                         then VT_VALUE else t(t(currentParentId).children(1)).nodeClass end;
      end if;
      
    end;
    
    procedure wrapExpression (sourceId in pls_integer) is
      childId   pls_integer;
      targetId  pls_integer := createPtg(null);
    begin
      t(targetId).children := t(sourceId).children;
      
      -- transfer source expression type to target expression 
      t(targetId).nodeType := t(sourceId).nodeType;
      t(targetId).nodeClass := t(sourceId).nodeClass;
      t(targetId).operandInfo := t(sourceId).operandInfo;
      t(sourceId).nodeType := null;
      
      for i in 1 .. t(targetId).children.count loop
        childId := t(targetId).children(i);
        t(childId).parentId := targetId;
      end loop;
      -- delete source children
      t(sourceId).children.delete;
      -- attach target to source
      pushOut(t(targetId));
    end;

    procedure pushExpr is
      nodeId  pls_integer;
    begin
      nodeId := createPtg(null);
      pushOut(t(nodeId));

      currentParentId := nodeId;      
    end;

    procedure pushSt (node in parseNode_t, createBranch boolean default true) is
    begin      
      st.extend;
      st(st.last) := node;
      
      -- push a new expression node and set as new parent node
      if createBranch then
        pushExpr();
      end if;
      
    end;
    
    function peek (list in parseNodeList_t, offset in pls_integer default 0) return parseNode_t is
    begin
      return list(list.last - offset);
    end;

    procedure pop (selectParent boolean default true) is
    begin
      st.trim;
      if selectParent then
        currentParentId := t(currentParentId).parentId;
      end if;
    end;
    
    function pop (selectParent boolean default true) return parseNode_t is
      top  parseNode_t := peek(st);
    begin
      pop(selectParent);
      return top;
    end;
    
    function isEmpty(list in parseNodeList_t) return boolean is
    begin
      return ( list is empty );
    end;

  begin
    
    if input is not null and input.count != 0 then
  
      rootId := createPtg(null);
      currentParentId := rootId;
      
      for i in 1 .. input.count loop
    
        node := t(input(i));
              
        case 
        when node.nodeType = XPTG_PARAM_SEP then
          
          loop
            if isEmpty(st) then
              error('Compilation error at position %d : separator misplaced or parentheses mismatched', 0 /*TODO*/);
            end if;
            top := peek(st);
            exit when isPtgFunc(top.nodeType);
            pop();
            pushOut(top);
          end loop;
          
          -- move parent one level up
          currentParentId := t(currentParentId).parentId;
          -- push new expression
          pushExpr();
          
        when opMetadataMap.exists(node.nodeType) then
          
          while not isEmpty(st) loop
            top := peek(st);
            if opMetadataMap.exists(top.nodeType)
               and (
                 ( opMetadataMap(node.nodeType).assoc = 0 and opMetadataMap(node.nodeType).prec <= opMetadataMap(top.nodeType).prec )
                 or opMetadataMap(node.nodeType).prec < opMetadataMap(top.nodeType).prec
               )
            then
              pop();
              pushOut(top);
              wrapExpression(currentParentId);
              
            else
              exit;
            end if;
          end loop;
          
          pushSt(node);
          
        when isPtgFunc(node.nodeType) then
          
          pushSt(node);
          
          
        when node.nodeType = XPTG_PAREN_LEFT then
          
          pushSt(node);
          
        when node.nodeType = XPTG_PAREN_RIGHT then
          
          loop
            if isEmpty(st) then
              error('Compilation error at position %d : parentheses mismatched', 0 /*TODO*/);
            end if;
            top := peek(st);
            exit when top.nodeType = XPTG_PAREN_LEFT;
            pop();
            pushOut(top);
          end loop;
          
          pop(); -- pop left parenthesis        
          pushOut(createPtgParen());
          wrapExpression(currentParentId);
          
        when node.nodeType = XPTG_FUNC_STOP then
          
          loop
            if isEmpty(st) then
              error('Compilation error at position %d : parentheses mismatched', 0 /*TODO*/);
            end if;
            top := peek(st);
            exit when isPtgFunc(top.nodeType);
            pop();
            pushOut(top);
          end loop; 
          
          pop();
          pushOut(top);
          
          ptgFuncVarList(top.id).argc := t(currentParentId).children.count - 1;
          
          if ptgFuncVarList(top.id).name = 'SUM' and ptgFuncVarList(top.id).argc = 1 then
            t(top.id).nodeType := PTG_ATTR_B;
            t(top.id).attrType := PTG_ATTRSUM;
            t(top.id).nodeClass := VT_NONE;
          end if;
          
          -- wrap function call in an expression
          wrapExpression(currentParentId);
          
        when node.nodeType in (PTG_STR, PTG_ERR, PTG_BOOL, PTG_INT, PTG_NUM, PTG_ARRAY_A, PTG_NAME_R, PTG_REF_R, PTG_AREA_R, PTG_NAMEX_R, PTG_REF3D_R, PTG_AREA3D_R, PTG_MISSARG) then
          
          pushOut(node);
          
        else
          
          error('Unexpected Ptg: %d', node.nodeType);
        
        end case;
      
      end loop;
      
      while not isEmpty(st) loop
        top := peek(st);
        if top.nodeType = XPTG_PAREN_LEFT or isPtgFunc(top.nodeType) then
          error('Compilation error at position %d : parentheses mismatched', 0 /*TODO*/);
        end if;
        pop();
        pushOut(top);
      end loop;
      
      -- if root node has a single child expression, make it the new root
      if t(rootId).children.count = 1 then
        firstChildId := t(rootId).children(1);
        --if isExpr(t(firstChildId)) then
          t(firstChildId).parentId := null;
          t.delete(rootId);
          rootId := firstChildId;
        --end if;
      end if;

      traverse(rootId);
      transformOperands(rootId);

    end if;
    
    return rootId;
  
  end;  


  function parseAll (
    input in varchar2
  )
  return pls_integer
  is
    token   token_t;
    area3d  area3d_t;
    name    name_t;
    
    idx     pls_integer;
    nodeId  pls_integer;
    nodeType  pls_integer;
    
    NULL_NODE  parseNode_t;
    
    nodeIdList  intList_t := intList_t();
    output      intList_t;
    
    stream  tokenStream_t;
    
    function getPtg (idx in pls_integer) return parseNode_t is
    begin
      if idx < 1 or idx > nodeIdList.last then
        return NULL_NODE;
      else
        return t(nodeIdList(idx));
      end if;
    end;
    
    function prevPtg (idx in pls_integer) return parseNode_t is
      prevIdx  pls_integer := nodeIdList.prior(idx);
    begin
      return case when prevIdx is not null then t(prevIdx) else NULL_NODE end;
    end;    

    function nextPtg (idx in pls_integer) return parseNode_t is
      nextIdx  pls_integer := nodeIdList.next(idx);
    begin
      return case when nextIdx is not null then t(nextIdx) else NULL_NODE end;
    end;
    
  begin
    
    if input is not null then
  
      null_node.id := -1;
      null_node.nodeType := 128;
      
      
      stream := tokenize(input);
      
      -- reset
      stream.idx := 0;
      t.delete;

      while hasNextToken(stream) loop
      
        token := getNextToken(stream);
        area3d := NULL_AREA3D;
        name := null;
      
        case 
        when getTokenTypeSequence(stream, 5) in (QUOTED_MULTISHEET_ROW_RANGE, QUOTED_MULTISHEET_COL_RANGE, QUOTED_MULTISHEET_CELL_RANGE) then
          area3d.prefix := nextToken(stream, 0).parsedValue.prefix;
          area3d.area.firstCell := nextToken(stream, 2).parsedValue.cell;
          area3d.area.lastCell := nextToken(stream, 4).parsedValue.cell;
          normalizePrefix(area3d.prefix);
          normalizeArea(area3d.area);
          stream.idx := stream.idx + 5 - 1;
          
          nodeId := createPtgArea3d(area3d, token.pos);
        
        when getTokenTypeSequence(stream, 3) = QUOTED_MULTISHEET_CELL_REF then
          area3d.prefix := nextToken(stream, 0).parsedValue.prefix;
          area3d.area.firstCell := nextToken(stream, 2).parsedValue.cell;
          normalizePrefix(area3d.prefix);
          stream.idx := stream.idx + 3 - 1;
          
          nodeId := createPtgRef3d(area3d, token.pos);
        
        when getTokenTypeSequence(stream, 7) in (MULTISHEET_ROW_RANGE, MULTISHEET_COL_RANGE, MULTISHEET_CELL_RANGE) then
          area3d.prefix.sheetRange.firstSheet := getSheet(nextToken(stream, 0).parsedValue.ident.value);
          area3d.prefix.sheetRange.lastSheet := getSheet(nextToken(stream, 2).parsedValue.ident.value);
          area3d.prefix.sheetRange.isSingleSheet := false;
          area3d.prefix.isNull := false;
          area3d.area.firstCell := nextToken(stream, 4).parsedValue.cell;
          area3d.area.lastCell := nextToken(stream, 6).parsedValue.cell;
          normalizePrefix(area3d.prefix);
          normalizeArea(area3d.area);
          stream.idx := stream.idx + 7 - 1;
          
          nodeId := createPtgArea3d(area3d, token.pos);
        
        when getTokenTypeSequence(stream, 5) = MULTISHEET_CELL_REF then
          area3d.prefix.sheetRange.firstSheet := getSheet(nextToken(stream, 0).parsedValue.ident.value);
          area3d.prefix.sheetRange.lastSheet := getSheet(nextToken(stream, 2).parsedValue.ident.value);
          area3d.prefix.sheetRange.isSingleSheet := false;
          area3d.prefix.isNull := false;
          area3d.area.firstCell := nextToken(stream, 4).parsedValue.cell;
          normalizePrefix(area3d.prefix);
          stream.idx := stream.idx + 5 - 1;
          
          nodeId := createPtgRef3d(area3d, token.pos);

        when getTokenTypeSequence(stream, 5) in (SHEET_ROW_RANGE, SHEET_COL_RANGE, SHEET_CELL_RANGE) then
          area3d.prefix.sheetRange.firstSheet := getSheet(nextToken(stream, 0).parsedValue.ident.value);
          area3d.prefix.isNull := false;
          area3d.area.firstCell := nextToken(stream, 2).parsedValue.cell;
          area3d.area.lastCell := nextToken(stream, 4).parsedValue.cell;
          normalizeArea(area3d.area);
          stream.idx := stream.idx + 5 - 1;
          
          nodeId := createPtgArea3d(area3d, token.pos);

        when getTokenTypeSequence(stream, 3) = SHEET_CELL_REF then
          area3d.prefix.sheetRange.firstSheet := getSheet(nextToken(stream, 0).parsedValue.ident.value);
          area3d.prefix.isNull := false;
          area3d.area.firstCell := nextToken(stream, 2).parsedValue.cell;     
          stream.idx := stream.idx + 3 - 1;
          
          nodeId := createPtgRef3d(area3d, token.pos);

        when getTokenTypeSequence(stream, 3) = SCOPED_NAME then
          name.scope := getSheet(nextToken(stream, 0).parsedValue.ident.value);
          name.value := nextToken(stream, 2).parsedValue.ident.value;
          stream.idx := stream.idx + 3 - 1;
          
          nodeId := createPtgNameX(name, token.pos);
          
        -- special case: defined name matches a sheet name 
        when getTokenTypeSequence(stream, 2) = SHEET_PREFIX and nextToken(stream, 2).type = T_NAME then
          name.scope := getSheet(nextToken(stream, 0).parsedValue.ident.value);
          name.value := nextToken(stream, 2).parsedValue.ident.value;
          stream.idx := stream.idx + 3 - 1;
          
          nodeId := createPtgNameX(name, token.pos);

        when getTokenTypeSequence(stream, 3) in (ROW_RANGE, COL_RANGE, CELL_RANGE) then
          area3d.area.firstCell := nextToken(stream, 0).parsedValue.cell;
          area3d.area.lastCell := nextToken(stream, 2).parsedValue.cell;
          normalizeArea(area3d.area);
          stream.idx := stream.idx + 3 - 1;
          
          nodeId := createPtgArea(area3d, token.pos);
        
        when token.type = T_CELL_REF then
          
          area3d.area.firstCell := token.parsedValue.cell;
          nodeId := createPtgRef(area3d, token.pos);


        -- fallback to T_NAME for unmatched T_COL_REF
        when token.type = T_COL_REF then
          
          if not token.parsedValue.cell.col.isAbsolute then
            name.value := token.parsedValue.cell.col.alphaValue;
            nodeId := createPtgName(name, token.pos);
          else
            error('Unexpected token: %s', tokenLabels(token.type));
          end if;

        when token.type = T_NAME then
          
          name.value := token.parsedValue.ident.value;
          nodeId := createPtgName(name, token.pos);
        
        when token.type = T_NUMBER then
          
          nodeId := createNumericPtg(token.value, token.pos);
          
        when token.type = T_STRING then
          
          nodeId := createPtgStr(token.value, token.pos);
            
        when token.type = T_ERROR then
          
          nodeId := createPtgErr(token.value, token.pos);

        when token.type = T_BOOLEAN then
          
          nodeId := createPtgBool(token.value, token.pos);
          
        when token.type = T_ARRAY_START then
          
          nodeId := createPtgArray(parseArray(stream), token.pos);
          
        when token.type = T_FUNC_START then
          
          nodeId := createPtgFuncVar(token.value, token.pos);
          
        when token.type = T_FUNC_STOP then
          
          nodeId := createPtg(XPTG_FUNC_STOP, token.value, token.pos);
          
        when token.type = T_LEFT then
          
          nodeId := createPtg(XPTG_PAREN_LEFT, pos => token.pos);

        when token.type = T_RIGHT then
          
          nodeId := createPtg(XPTG_PAREN_RIGHT, pos => token.pos);

        when token.type = T_WSPACE then
          
          nodeId := createPtg(XPTG_SPACE, pos => token.pos);
          
        when token.type = T_ARG_SEP then
          
          nodeId := createPtg(XPTG_PARAM_SEP, pos => token.pos);
          
        when token.type between OP_PLUS and OP_PERCENT then
          
          nodeType := tokenOpPtgMap(token.type);
          nodeId := createPtg(nodeType, opMetadataMap(nodeType).token, token.pos);
        
        when token.type = T_MISSING_ARG then
          
          nodeId := createPtg(PTG_MISSARG, pos => token.pos);
        
        else
          
          error('Unexpected token: %s', tokenLabels(token.type));
        
        end case;
        
        nodeIdList.extend;
        nodeIdList(nodeIdList.last) := nodeId;
      
      end loop;

      for i in 1 .. nodeIdList.count loop
        
        -- convert space between two cell references to intersection operator 
        if getPtg(i).nodeType = XPTG_SPACE then
          if isRef(prevPtg(i), MATCH_FUNC_STOP) and isRef(nextPtg(i), MATCH_FUNC_START) 
             -- also match a space occurring between a ref and an opening/closing parenthesis, or two parentheses, e.g. SUM((A1 A1) A1)
             -- we'll sort this situation out later in the compile stage in case the bracketed expression is not an actual ref
             or isRef(prevPtg(i), MATCH_FUNC_STOP) and nextPtg(i).nodeType = XPTG_PAREN_LEFT
             or prevPtg(i).nodeType = XPTG_PAREN_RIGHT and isRef(nextPtg(i), MATCH_FUNC_START)
             or prevPtg(i).nodeType = XPTG_PAREN_RIGHT and nextPtg(i).nodeType = XPTG_PAREN_LEFT
          then
            t(nodeIdList(i)).nodeType := PTG_ISECT;
            t(nodeIdList(i)).token := opMetadataMap(PTG_ISECT).token;
            --t(nodeIdList(i)).nodeClass := opMetadataMap(PTG_ISECT).returnClass;
          elsif prevPtg(i).nodeType in (PTG_STR, PTG_ERR, PTG_BOOL, PTG_INT, PTG_NUM) 
             and not opMetadataMap.exists(nextPtg(i).nodeType)
             or nextPtg(i).nodeType in (PTG_STR, PTG_ERR, PTG_BOOL, PTG_INT, PTG_NUM) 
             and not opMetadataMap.exists(prevPtg(i).nodeType)
            
          then
            error('Unexpected token: %s', ptgLabels(XPTG_SPACE)); -- TODO: at position %d
          else
            -- discard this node
            nodeIdList.delete(i);
          end if;
        end if;
          
      end loop;
      
      -- nodeIdList may have become sparse at this point, switching to iterator method    
      idx := nodeIdList.first;   
      while idx is not null loop
        
        -- convert minus to unary minus where needed 
        if getPtg(idx).nodeType = PTG_SUB then
          
          if not(isRef(prevPtg(idx), MATCH_FUNC_STOP) 
                 or prevPtg(idx).nodeType in (XPTG_FUNC_STOP, XPTG_PAREN_RIGHT, PTG_ARRAY_A, PTG_INT, PTG_NUM, PTG_STR, PTG_BOOL, PTG_ERR, PTG_PERCENT)) 
          then
            case nextPtg(idx).nodeType
            when PTG_INT then
              -- convert to negative double, discard minus
              t(nodeIdList.next(idx)).nodeType := PTG_NUM;
              t(nodeIdList.next(idx)).token := '-' || t(nodeIdList.next(idx)).token;
              ptgNumList(nodeIdList.next(idx)).value := to_binary_double(-1 * ptgIntList(nodeIdList.next(idx)).value);
              ptgIntList.delete(nodeIdList.next(idx));
              nodeIdList.delete(idx);
            when PTG_NUM then
              -- apply sign and discard minus
              t(nodeIdList.next(idx)).token := '-' || t(nodeIdList.next(idx)).token;
              ptgNumList(nodeIdList.next(idx)).value := -1 * ptgNumList(nodeIdList.next(idx)).value;
              nodeIdList.delete(idx);
            else
              -- convert to unary minus
              t(nodeIdList(idx)).nodeType := PTG_UMINUS;
            end case;
          end if;
          
        -- discard unary plus
        elsif getPtg(idx).nodeType = PTG_ADD then
        
          if not(isRef(prevPtg(idx), MATCH_FUNC_STOP) 
                 or prevPtg(idx).nodeType in (XPTG_FUNC_STOP, XPTG_PAREN_RIGHT, PTG_ARRAY_A, PTG_INT, PTG_NUM, PTG_STR, PTG_BOOL, PTG_ERR, PTG_PERCENT)) 
          then
            nodeIdList.delete(idx);
          end if;      
        
        end if;
        
        idx := nodeIdList.next(idx);

      end loop;

      -- copy to a dense list
      output := intList_t();
      idx := nodeIdList.first;   
      while idx is not null loop
        output.extend;
        output(output.last) := nodeIdList(idx);
        idx := nodeIdList.next(idx);
      end loop;

      /*
      for i in 1 .. output.count loop
        dbms_output.put_line(
          utl_lms.format_message(
            '%s'
          , ptgLabels(t(output(i)).nodeType)
          )
        );
      end loop;
      */
    
    end if;
    
    return compile(output);
    
  end;
  
  function getParseTree (input in varchar2) return parseTree_t pipelined is
    
    rootId  pls_integer := parseAll(input);
    
    type miniStackEntry_t is record (nodeId pls_integer, childIdx pls_integer);
    type miniStack_t is table of miniStackEntry_t;
    
    st  miniStack_t := miniStack_t();

    nodeId     pls_integer;
    node       parseNode_t;
    nodeLevel  pls_integer := 1;
    currentChildIndex  pls_integer;
    
    procedure push is
    begin
      currentChildIndex := 1;
      st.extend;
      st(st.last).childIdx := currentChildIndex;
    end;
    
    function makeTreeNode return parseTreeNode_t is
      treeNode  parseTreeNode_t;
    begin
      treeNode.id := node.id;
      treeNode.nodeType := node.nodeType;
      treeNode.nodeTypeName := ptgLabel(node);
      treeNode.parentId := node.parentId;
      treeNode.nodeLevel := nodeLevel;
      treeNode.token := node.token;
      treeNode.nodeClass := node.nodeClass;
      return treeNode; 
    end;
    
  begin
      
    ctx.treeRootId := rootId;
    
    if rootId is not null then
    
      nodeId := rootId;
    
      <<main>>
      loop
        
        node := t(nodeId);
        pipe row ( makeTreeNode() );
        
        if node.children is not empty then
        
          push();
          nodeId := node.children(currentChildIndex);
          nodeLevel := nodeLevel + 1;
        
        else
          
          if node.parentId is not null then
            
            loop
              
              node := t(node.parentId);
              
              -- get next sibling
              if currentChildIndex < node.children.count then
                currentChildIndex := currentChildIndex + 1;
                st(st.last).childIdx := currentChildIndex;
                nodeId := node.children(currentChildIndex);
                exit;
              else
                -- parent has no more children to visit
                nodeLevel := nodeLevel - 1;
                st.trim;
                
                if nodeLevel = 1 then
                  -- we are at the root node : exit all
                  exit main;
                end if;
                
                currentChildIndex := st(st.last).childIdx;
                
              end if;
            
            end loop;
            
          else
            exit main;
          end if;
        
        end if;
        
      end loop;
    
    end if;
    
    return;
  
  end;

  function getFunctionalTree (input in varchar2) return parseTree_t pipelined is
    
    rootId  pls_integer;
    
    type miniStackEntry_t is record (nodeId pls_integer, childIdx pls_integer);
    type miniStack_t is table of miniStackEntry_t;
    
    st  miniStack_t := miniStack_t();

    nodeId     pls_integer;
    node       parseNode_t;
    nodeLevel  pls_integer := 1;
    currentChildIndex  pls_integer;
    lastChildId  pls_integer;
    
    procedure push is
    begin
      currentChildIndex := 1;
      st.extend;
      st(st.last).childIdx := currentChildIndex;
    end;
    
    function makeTreeNode return parseTreeNode_t is
      treeNode  parseTreeNode_t;
    begin
      treeNode.id := node.id;
      treeNode.nodeType := node.nodeType;
      treeNode.nodeTypeName := ptgLabel(node);
      treeNode.parentId := node.parentId;
      treeNode.nodeLevel := nodeLevel;
      treeNode.token := node.token;
      treeNode.nodeClass := node.nodeClass;
      treeNode.nodeStatus := case when node.children is empty then 0 else 1 end;
      return treeNode; 
    end;
    
  begin
    
    begin
      
    rootId := parseAll(input);
  
    if rootId is not null then
  
      nodeId := rootId;
    
      <<main>>
      loop
        
        node := t(nodeId);
        
        if node.nodeType = EXPR_MEM_AREA then
          lastChildId := node.children(node.children.last);
          t(lastChildId).parentId := node.parentId;
          if node.parentId is not null then
            node.children(currentChildIndex) := lastChildId;
          end if;
          node := t(lastChildId);
        end if;
        
        if isExpr(node) then
          -- last child represents either an operator or a function
          lastChildId := node.children(node.children.last);
          node.children.trim;
          -- substitute current node with that child in the tree
          t(lastChildId).parentId := node.parentId;
          -- attach operands to their new parent
          for i in 1 .. node.children.count loop
            t(node.children(i)).parentId := lastChildId;
          end loop;
          -- adopt children
          t(lastChildId).children := node.children;
          t.delete(nodeId);
          node := t(lastChildId);
        end if;
        
        pipe row ( makeTreeNode() );
        
        if node.children is not empty then
        
          push();
          nodeId := node.children(currentChildIndex);
          nodeLevel := nodeLevel + 1;
        
        else
          
          if node.parentId is not null then
          
            loop
              
              node := t(node.parentId);
              
              -- get next sibling
              if currentChildIndex < node.children.count then
                currentChildIndex := currentChildIndex + 1;
                st(st.last).childIdx := currentChildIndex;
                nodeId := node.children(currentChildIndex);
                exit;
              else
                -- parent has no more children to visit
                nodeLevel := nodeLevel - 1;
                st.trim;
                
                if nodeLevel = 1 then
                  -- we are at the root node : exit all
                  exit main;
                end if;
                
                currentChildIndex := st(st.last).childIdx;
                
              end if;
            
            end loop;
            
          else
            exit main;
          end if;
        
        end if;
        
      end loop;

    end if;
    
    exception
      when others then
        null;
    
    end;
       
    return;
  
  end;

  function getNodeList (rootId in pls_integer) return parseNodeList_t is
    
    nl  parseNodeList_t := t;
    
    type miniStackEntry_t is record (nodeId pls_integer, childIdx pls_integer);
    type miniStack_t is table of miniStackEntry_t;
    
    st  miniStack_t := miniStack_t();

    nodeId     pls_integer;
    node       parseNode_t;
    nodeLevel  pls_integer := 1;
    currentChildIndex  pls_integer;
    lastChildId  pls_integer;
    
    procedure push is
    begin
      currentChildIndex := 1;
      st.extend;
      st(st.last).childIdx := currentChildIndex;
    end;
    
    function makeTreeNode return parseTreeNode_t is
      treeNode  parseTreeNode_t;
    begin
      treeNode.id := node.id;
      treeNode.nodeType := node.nodeType;
      treeNode.nodeTypeName := ptgLabel(node);
      treeNode.parentId := node.parentId;
      treeNode.nodeLevel := nodeLevel;
      treeNode.token := node.token;
      treeNode.nodeClass := node.nodeClass;
      treeNode.nodeStatus := case when node.children is empty then 0 else 1 end;
      return treeNode; 
    end;
    
  begin
    
    if rootId is not null then
  
      nodeId := rootId;
    
      <<main>>
      loop
        
        node := nl(nodeId);
        
        if isExpr(node) then
          -- last child represents either an operator or a function
          lastChildId := node.children(node.children.last);
          node.children.trim;
          -- substitute current node with that child in the tree
          nl(lastChildId).parentId := node.parentId;
          -- attach operands to their new parent
          for i in 1 .. node.children.count loop
            nl(node.children(i)).parentId := lastChildId;
          end loop;
          -- adopt children
          nl(lastChildId).children := node.children;
          nl.delete(nodeId);
          node := nl(lastChildId);
        end if;
        
        if node.children is not empty then
        
          push();
          nodeId := node.children(currentChildIndex);
          nodeLevel := nodeLevel + 1;
        
        else
          
          if node.parentId is not null then
          
            loop
              
              node := t(node.parentId);
              
              -- get next sibling
              if currentChildIndex < node.children.count then
                currentChildIndex := currentChildIndex + 1;
                st(st.last).childIdx := currentChildIndex;
                nodeId := node.children(currentChildIndex);
                exit;
              else
                -- parent has no more children to visit
                nodeLevel := nodeLevel - 1;
                st.trim;
                
                if nodeLevel = 1 then
                  -- we are at the root node : exit all
                  exit main;
                end if;
                
                currentChildIndex := st(st.last).childIdx;
                
              end if;
            
            end loop;
            
          else
            exit main;
          end if;
        
        end if;
        
      end loop;
    
    end if;
    
    return null;
  
  end;
  
  function toBin (value in pls_integer, sz in pls_integer default 4) return raw is
  begin
    return utl_raw.substr(utl_raw.cast_from_binary_integer(value, utl_raw.little_endian), 1, sz);
  end;
  
  -- 2.5.153 UncheckedCol
  function uncheckedCol (col in column_t) return raw is
  begin
    return toBin(col.value - 1);
  end;
  
  -- 2.5.155 UncheckedRw
  function uncheckedRw (rw in row_t) return raw is
  begin
    return toBin(rw.value - 1);
  end;
  
  -- 2.5.129 RwRelNeg
  function rwRelNeg (rw in row_t) return raw is
  begin
    return toBin(rw.value - 1);
  end;
  
  -- 2.5.26 ColRelShort
  function colRelShort (cell in cell_t) return raw is
  begin
    return toBin(
               cell.col.value - 1
             + case when cell.col.isAbsolute then 0 else 16384 end -- fColRel (bit 14)
             + case when cell.rw.isAbsolute then 0 else 32768 end -- fRwRel (bit 15)
           , 2);
  end;

  -- 2.2.7 External References
  function putExternal (
    supLink     in pls_integer
  , firstSheet  in pls_integer
  , lastSheet   in pls_integer default null
  )
  return pls_integer
  is
    xti     ExcelTypes.xti_t;
    xtiKey  varchar2(24);
  begin
    
    -- 2.2.7.2 Supporting Link
    -- zero-based index of this supporting link in the collection of supporting links
    if ctx.externals.supLinks.exists(supLink) then
      xti.externalLink := ctx.externals.supLinks(supLink);
    else
      xti.externalLink := ctx.externals.supLinks.count;
      ctx.externals.supLinks(supLink) := xti.externalLink;
    end if;

    -- 2.5.173 Xti
    xti.firstSheet := firstSheet;
    xti.lastSheet := nvl(lastSheet, firstSheet);
    xtiKey := rawtohex(utl_raw.concat(toBin(xti.externalLink), toBin(xti.firstSheet), toBin(xti.lastSheet)));
    
    if ctx.externals.xtiMap.exists(xtiKey) then
      xti := ctx.externals.xtiMap(xtiKey);
    else
      ctx.externals.xtiArray.extend;
      xti.idx := ctx.externals.xtiArray.last - 1;
      ctx.externals.xtiArray(xti.idx + 1) := xti;
      ctx.externals.xtiMap(xtiKey) := xti;
    end if;
    
    -- 2.5.98.103 XtiIndex
    return xti.idx;
    
  end;
  
  procedure putPtgType (ptg in parseNode_t) is
  begin
    putBytes(toBin(ptg.nodeType, 1));
  end;

  procedure putStr (str in varchar2) is
  begin
    putBytes(utl_raw.concat(toBin(length(str),2), utl_i18n.string_to_raw(str, 'AL16UTF16LE')));
  end;
  
  -- 2.5.98.23 PtgArray
  procedure putPtgArray (arr in ptgArray_t) is
    item  arrayItem_t;
  begin
    putBytes( utl_raw.concat(
             '00000000'  -- unused1
           , '0000'      -- unused2
           , '00000000'  -- unused3
           , '00000000'  -- unused4
           ));
    
    -- 2.5.98.41 PtgExtraArray
    putExtra(toBin(arr.value.count)); -- rows
    putExtra(toBin(arr.value(1).count)); -- cols
    for r in 1 .. arr.value.count loop
      for c in 1 .. arr.value(r).count loop
        item := arr.value(r)(c);
        putExtra(toBin(item.itemType, 1)); -- reserved byte (item type)
        case item.itemType
        when 0 then -- SerNum
          putExtra(utl_raw.cast_from_binary_double(item.numValue, utl_raw.little_endian));
        when 1 then -- SerStr
          putExtra(utl_raw.concat(toBin(length(item.strValue)), utl_i18n.string_to_raw(item.strValue, 'AL16UTF16LE')));
        when 2 then -- SerBool
          putExtra(toBin(item.boolValue, 1));
        when 4 then -- SerErr
          putExtra(toBin(item.errValue, 1));
          putExtra('000000'); -- reserved2, reserved3
        end case;
      end loop;
    end loop;
    
  end;
  
  procedure putPtgFunc (func in ptgFuncVar_t) is
  begin
    putBytes(toBin(functionMetadataMap(func.name).id, 2));
  end;

  procedure putPtgFuncVar (func in ptgFuncVar_t) is
  begin
    putBytes(toBin(func.argc, 1)); -- cparams
    putBytes(toBin(functionMetadataMap(func.name).id, 2));
  end;
  
  procedure putRgceLoc (cell in cell_t) is
  begin
    putBytes(uncheckedRw(cell.rw));
    putBytes(colRelShort(cell));
  end;

  procedure putRgceLocRel (cell in cell_t) is
  begin
    -- row
    if cell.rw.isAbsolute then
      putBytes(uncheckedRw(cell.rw));
    else
      putBytes(rwRelNeg(cell.rw));
    end if;
    putBytes(colRelShort(cell));
  end;
  
  procedure putRgceArea (area in area_t) is
  begin
    putBytes(uncheckedRw(area.firstCell.rw)); -- rowFirst
    putBytes(uncheckedRw(area.lastCell.rw));  -- rowLast
    putBytes(colRelShort(area.firstCell));    -- columnFirst
    putBytes(colRelShort(area.lastCell));     -- columnLast
  end;

  procedure putRgceAreaRel (area in area_t) is
  begin
    -- rowFirst
    if area.firstCell.rw.isAbsolute then
      putBytes(uncheckedRw(area.firstCell.rw));
    else
      putBytes(rwRelNeg(area.firstCell.rw));
    end if;
    -- rowLast
    if area.lastCell.rw.isAbsolute then
      putBytes(uncheckedRw(area.lastCell.rw));
    else
      putBytes(rwRelNeg(area.lastCell.rw));
    end if;    
    putBytes(colRelShort(area.firstCell));  -- columnFirst
    putBytes(colRelShort(area.lastCell));   -- columnLast
  end;
  
  procedure putPtg (nodeId in pls_integer);

  procedure putChildren (ptg in parseNode_t) is
  begin
    for i in 1 .. ptg.children.count loop
      putPtg(ptg.children(i));
    end loop;    
  end;
  
  procedure putMemAreaExpr (expr in parseNode_t) is
    memPtg    parseNode_t := t(expr.children(1));
    areaList  areaList_t;
    ccePtr    pls_integer;
  begin
    
    putPtgType(memPtg);
  
    case
    when memPtg.nodeType in (PTG_MEMAREA_R, PTG_MEMAREA_V, PTG_MEMAREA_A) then
      putBytes('00000000'); -- unused
      
      -- 2.5.98.44 PtgExtraMem
      areaList := ptgMemAreaList(memPtg.id).value;
      putExtra(toBin(areaList.count));
      for i in 1 .. areaList.count loop
        -- 2.5.154 UncheckedRfX
        putExtra(uncheckedRw(areaList(i).firstCell.rw));   -- rwFirst
        putExtra(uncheckedRw(areaList(i).lastCell.rw));    -- rwLAst
        putExtra(uncheckedCol(areaList(i).firstCell.col)); -- colFirst
        putExtra(uncheckedCol(areaList(i).lastCell.col));  -- colLast
      end loop;
    
    when memPtg.nodeType in (PTG_MEMFUNC_R, PTG_MEMFUNC_V, PTG_MEMFUNC_V) then
      null;
      
    when memPtg.nodeType in (PTG_MEMERR_R, PTG_MEMERR_V, PTG_MEMERR_A) then
      putBytes(toBin(ptgMemErrList(memPtg.id).value, 1)); -- err
      putBytes('000000'); -- unused1 + unused2
    
    else
      error('Unexpected Ptg: %s', ptgLabel(memPtg));
    end case;
    
    ccePtr := currentStream.sz + 1;
    putBytes('0000'); -- cce field initialized to zero
    
    putPtg(expr.children(2)); -- binary reference expression
    putBytes(toBin(currentStream.sz - (ccePtr + 2) + 1, 2), ccePtr); -- cce: size of the binary ref expr
    
  end;
  
  procedure putIfExpr (ptg in parseNode_t) is
    ptgAttrIfPtr  pls_integer;
    gotos         intList_t := intList_t(); -- list of pointers to PtgAttrGoTo
  begin
    putPtg(ptg.children(1)); -- arg1 (conditional expr)
    
    -- 2.5.98.27 PtgAttrIf
    ptgAttrIfPtr := currentStream.sz + 1;
    putBytes('19020000'); -- offset initialized to zero
    
    putPtg(ptg.children(2)); -- arg2
    
    -- 2.5.98.26 PtgAttrGoTo
    gotos.extend;
    gotos(gotos.last) := currentStream.sz + 1;
    putBytes('19080000'); -- offset initialized to zero
    
    -- Set PtgAttrIf offset field
    putBytes(toBin(currentStream.sz + 1 - (ptgAttrIfPtr + 4), 2), ptgAttrIfPtr + 2);
    
    -- arg3?
    if ptg.children.count > 3 then    
      putPtg(ptg.children(3));
      -- PtgAttrGoTo
      gotos.extend;
      gotos(gotos.last) := currentStream.sz + 1;
      putBytes('19080000');
    end if;
        
    putPtgFuncVar(ptgFuncVarList(ptg.id));

    for i in 1 .. gotos.count loop
      putBytes(toBin(currentStream.sz - gotos(i) - 4, 2), gotos(i)+2);
    end loop;
    
  end;

  procedure putChooseExpr (ptg in parseNode_t) is
    choiceCount   pls_integer := ptg.children.count - 2;
    jumpTablePtr  pls_integer;
    choicePtr     pls_integer;
    gotos         intList_t := intList_t(); -- list of pointers to PtgAttrGoTo
  begin
    putPtg(ptg.children(1)); -- index arg
    
    -- 2.5.98.25 PtgAttrChoose
    putBytes('1904');
    putBytes(toBin(choiceCount - 1, 2)); -- cOffset
    jumpTablePtr := currentStream.sz + 1;
    putBytes(utl_raw.copies('00', choiceCount * 2)); -- rgOffset: array of 2-byte integers initialized to zero
    
    for i in 1 .. choiceCount loop
      
      choicePtr := currentStream.sz + 1;
      putBytes(toBin(choicePtr - jumpTablePtr, 2), jumpTablePtr + 2*(i - 1)); -- set entry in rgOffset
      putPtg(ptg.children(i+1)); -- choice #i
      -- PtgAttrGoTo
      gotos.extend;
      gotos(gotos.last) := currentStream.sz + 1; -- store ptr of this PtgAttrGoTo
      putBytes('19080000'); -- set offset to zero, to be filled later
    
    end loop;
    
    putPtgFuncVar(ptgFuncVarList(ptg.id));
    
    -- Set PtgAttrGoTo's offset field
    -- Each offset is the size of all remaining Ptg's after this PtgAttrGoTo, minus one.
    -- It is equivalent to subtracting the position right after this PtgAttrGoTo from the position right after the terminating PtgFuncVar
    for i in 1 .. gotos.count loop
      putBytes(toBin(currentStream.sz - gotos(i) - 4, 2), gotos(i)+2);
    end loop;
    
  end;

  -- 2.5.98.60 PtgName
  procedure putPtgName (name in name_t) is
  begin
    putBytes(toBin(name.idx)); -- nameindex
  end;

  -- 2.5.98.61 PtgNameX
  procedure putPtgNameX (name in name_t) is
    ixti  pls_integer := putExternal(357 /*BrtSupSelf*/, -2);
  begin
    putBytes(toBin(ixti, 2));   -- ixti
    putBytes(toBin(name.idx));  -- nameindex
  end;
  
  procedure put_xlfnName (nm in name_t)
  is
    definedName  ExcelTypes.CT_DefinedName;
  begin
    ctx.definedNames.extend;
    
    definedName.idx := nm.idx;
    definedName.name := nm.value;
    definedName.formula := '#NAME?';
    definedName.hidden := true;
    definedName.futureFunction := true;
    
    ctx.definedNames(ctx.definedNames.last) := definedName;
    
  end;
  
  procedure putFuncCall (ptg in parseNode_t) is
    funcPtgId  pls_integer := ptg.children(ptg.children.last);
    funcName   varchar2(256) := ptgFuncVarList(funcPtgId).name;
    funcMeta   functionMetadata_t;
    xlfnName   name_t;
    nodeType   pls_integer;
  begin
    case funcName
    when 'IF' then
      putIfExpr(ptg);
    when 'CHOOSE' then
      putChooseExpr(ptg);
    else
      
      -- 2.5.98.10 Ftab #future-function
      -- TODO? push a PtgName earlier in the parse tree so that it is handled like the others by putPtg.
      funcMeta := functionMetadataMap(funcName);
      if funcMeta.id = 255 and funcMeta.internalName like '_xlfn.%' then
        
        if ctx.nameMap.exists(upper(funcMeta.internalName)) then
          xlfnName := ctx.nameMap(upper(funcMeta.internalName));
        else
          xlfnName := putName(funcMeta.internalName);
          put_xlfnName(xlfnName);
        end if;
        -- PtgName
        putBytes(toBin(PTG_NAME_R, 1));
        putPtgName(xlfnName);
        
        -- Apparently, when a future-function is used, the Ptg becomes a PtgFuncVar regardless of the original param count,
        -- and that count is incremented by one to take the leading PtgName into account
        nodeType := PTG_FUNCVAR_R;
        applyPtgDataType(nodeType, t(funcPtgId).nodeClass);
        t(funcPtgId).nodeType := nodeType;
        ptgFuncVarList(funcPtgId).argc := ptgFuncVarList(funcPtgId).argc + 1;
        
      end if;
    
      putChildren(ptg);
    end case;
  end;
  
  procedure putRef3d (ref3d in area3d_t) is
    ixti  pls_integer := putExternal(357, ref3d.prefix.sheetRange.firstSheet.idx, ref3d.prefix.sheetRange.lastSheet.idx);
  begin
    putBytes(toBin(ixti, 2));
    -- loc
    if ctx.config.fmlaType = EXC_FMLATYPE_NAME then
      putRgceLocRel(ref3d.area.firstCell);
    else
      putRgceLoc(ref3d.area.firstCell);
    end if;
  end;

  procedure putArea3d (area3d in area3d_t) is
    ixti  pls_integer := putExternal(357, area3d.prefix.sheetRange.firstSheet.idx, area3d.prefix.sheetRange.lastSheet.idx);
  begin
    putBytes(toBin(ixti, 2));
    -- area
    if ctx.config.fmlaType = EXC_FMLATYPE_NAME then
      putRgceAreaRel(area3d.area);
    else
      putRgceArea(area3d.area);
    end if;
  end;
  
  procedure putPtg (nodeId in pls_integer) is
    ptg  parseNode_t := t(nodeId);
  begin

    if isExpr(ptg) then
      
      case ptg.nodeType
      when EXPR_MEM_AREA then
      
        putMemAreaExpr(ptg);
        
      when EXPR_FUNC_CALL then
        
        putFuncCall(ptg);
        
      else
        
        putChildren(ptg);
        
      end case;
      
    elsif isBasePtg(ptg.nodeType) then

      putPtgType(ptg);

      case ptg.nodeType
      when PTG_STR then
        putStr(ptgStrList(ptg.id).value);
        
      when PTG_ERR then
        putBytes(toBin(ptgErrList(ptg.id).value, 1));
        
      when PTG_BOOL then
        putBytes(toBin(ptgBoolList(ptg.id).value, 1));
        
      when PTG_INT then
        putBytes(toBin(ptgIntList(ptg.id).value, 2));
        
      when PTG_NUM then
        putBytes(utl_raw.cast_from_binary_double(ptgNumList(ptg.id).value, utl_raw.little_endian));
        
      when PTG_ATTR_B then
        if ptg.attrType = PTG_ATTRSUM then
          putBytes(toBin(ptg.attrType, 1));
          putBytes('0000'); -- unused
        else
          error('Unsupported PtgAttr: %d', ptg.attrType);
        end if;
        
      else
        -- should be an operator
        null;
      
      end case;
      
    else
      
      putPtgType(ptg);
    
      case
      when ptg.nodeType in (PTG_ARRAY_R, PTG_ARRAY_V, PTG_ARRAY_A) then
        putPtgArray(ptgArrayList(ptg.id));
        
      when ptg.nodeType in (PTG_FUNC_R, PTG_FUNC_V, PTG_FUNC_A) then
        putPtgFunc(ptgFuncVarList(ptg.id));
      
      when ptg.nodeType in (PTG_FUNCVAR_R, PTG_FUNCVAR_V, PTG_FUNCVAR_A) then
        putPtgFuncVar(ptgFuncVarList(ptg.id));
        
      when ptg.nodeType in (PTG_NAME_R, PTG_NAME_V, PTG_NAME_A) then
        putPtgName(ptgNameList(ptg.id).value);
      
      when ptg.nodeType in (PTG_NAMEX_R, PTG_NAMEX_V, PTG_NAMEX_A) then
        putPtgNameX(ptgNameList(ptg.id).value);
        
      when ptg.nodeType in (PTG_REF_R, PTG_REF_V, PTG_REF_A) then
        putRgceLoc(ptgArea3dList(ptg.id).value.area.firstCell);
        
      when ptg.nodeType in (PTG_AREA_R, PTG_AREA_V, PTG_AREA_A) then
        putRgceArea(ptgArea3dList(ptg.id).value.area);
        
      when ptg.nodeType in (PTG_REF3D_R, PTG_REF3D_V, PTG_REF3D_A) then
        putRef3d(ptgArea3dList(ptg.id).value);
      
      when ptg.nodeType in (PTG_AREA3D_R, PTG_AREA3D_V, PTG_AREA3D_A) then
        putArea3d(ptgArea3dList(ptg.id).value);
      
      end case;
      
    
    end if;
    
    
  end;
  
  function serialize (nodeId in pls_integer)
  return stream_t  
  is
    ptg        parseNode_t := t(nodeId);
    funcPtgId  pls_integer;
    str        stream_t;
  begin

    if ptg.nodeType = EXPR_MEM_AREA then
      
      putChars(str, serialize(ptg.children(2)));
      
    elsif ptg.nodeType = EXPR_FUNC_CALL then
      
      funcPtgId := ptg.children(ptg.children.last);
      putChars(str, t(funcPtgId).token);
      putChars(str, '(');
      for i in 1 .. ptg.children.count - 1 loop
        if i > 1 then
          putChars(str, ',');
        end if;
        putChars(str, serialize(ptg.children(i)));
      end loop;
      putChars(str, ')');
      
    elsif ptg.nodeType in (EXPR_BINARY_REF, EXPR_BINARY_VAL) then

      putChars(str, serialize(ptg.children(1)));
      putChars(str, t(ptg.children(3)).token);
      putChars(str, serialize(ptg.children(2)));

    elsif ptg.nodeType = EXPR_DISPLAY_PREC then
      
      putChars(str, '(');
      putChars(str, serialize(ptg.children(1)));
      putChars(str, ')');
    
    elsif ptg.nodeType = EXPR_UNARY then
      
      putChars(str, t(ptg.children(2)).token);
      putChars(str, serialize(ptg.children(1)));
    
    else
      
      putChars(str, ptg.token);
    
    end if;
    
    return str;
        
  end;
  
  procedure setContext (
    p_sheets  in ExcelTypes.CT_Sheets
  , p_names   in ExcelTypes.CT_DefinedNames
  )
  is
  begin
    
    ctx.externals.supLinks.delete;
    ctx.externals.xtiMap.delete;
    ctx.externals.xtiArray := ExcelTypes.xtiArray_t();
    
    ctx.sheetMap.delete;
    for i in 1 .. p_sheets.count loop
      putSheet(p_sheets(i).name, p_sheets(i).idx);
    end loop;
    
    ctx.nameMap.delete;
    for i in 1 .. p_names.count loop
      putName(p_names(i).name, p_names(i).scope, p_names(i).idx);
    end loop;
    
  end;
  
  function getFormula return varchar2
  is
  begin
    return serialize(ctx.treeRootId).chars.content;
  end;
  
  function parse (p_expr in varchar2) return varchar2
  is
  begin
    return serialize(parseAll(p_expr)).chars.content;
  end;
  
  function parseBinary (
    p_expr     in varchar2
  , p_type     in pls_integer default null
  , p_cellRef  in varchar2 default null
  ) 
  return raw
  is
    rootId  pls_integer;
  begin
    setFormulaType(nvl(p_type, EXC_FMLATYPE_CELL));
    setCurrentCell(nvl(p_cellRef, 'A1'));
    ctx.definedNames := ExcelTypes.CT_DefinedNames();
    ctx.volatile := false;
    rootId := parseAll(p_expr);
    initStream(currentStream);
    
    if ctx.volatile then
      -- 2.5.98.29 PtgAttrSemi
      putBytes(toBin(PTG_ATTR_B, 1));
      putBytes(toBin(PTG_ATTRSEMI, 1));
      putBytes('0000'); -- unused
    end if;
    
    putPtg(rootId);
    --TODO: pass back generated names
    return utl_raw.concat(toBin(currentStream.sz)      -- cce
                        , currentStream.bytes.content  -- rgce
                        , toBin(ctx.rgbExtra.sz)       -- cb
                        , ctx.rgbExtra.bytes.content   -- rgcb
                        );
  end;
  
  function getExternals return ExcelTypes.CT_Externals
  is
  begin
    return ctx.externals;
  end;
  
  function getNames return ExcelTypes.CT_DefinedNames
  is
  begin
    return ctx.definedNames;
  end;
   
begin
  
  initState;

end ExcelFmla;
/
