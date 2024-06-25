create or replace package ExcelFmla is
/* ======================================================================================

  This Source Code Form is subject to the terms of the Mozilla Public 
  License, v. 2.0. If a copy of the MPL was not distributed with this 
  file, You can obtain one at https://mozilla.org/MPL/2.0/.

  Copyright (c) 2023-2024 Marc Bleron  

=========================================================================================
    Change history :
    Marc Bleron       2023-10-01     Creation
====================================================================================== */

  FMLATYPE_CELL    constant pls_integer := 0;
  FMLATYPE_SHARED  constant pls_integer := 2;
  FMLATYPE_NAME    constant pls_integer := 5;
  
  REF_R1C1         constant pls_integer := 0;
  REF_A1           constant pls_integer := 1;
  
  type parseTreeNode_t is record (
    id            integer
  , nodeType      integer
  , nodeTypeName  varchar2(128)
  , parentId      integer
  , nodeLevel     integer
  , token         varchar2(32767)
  , nodeClass     integer
  , nodeStatus    integer
  );
  
  type parseTree_t is table of parseTreeNode_t;

  subtype sheet_name is varchar2(31 char);
  type sheetList_t is table of sheet_name;
  
  procedure setDebug (enabled in boolean default true);

  function base26Decode (input in varchar2) return pls_integer result_cache;
  function base26Encode (input in pls_integer) return varchar2 result_cache;
  function isValidCellReference (input in varchar2, refStyle in pls_integer default REF_A1) return boolean;
  
  procedure putName (value in varchar2, sheetName in varchar2 default null, idx in pls_integer default null);
  procedure putSheet (name in varchar2, idx in pls_integer default null);
  procedure setCurrentSheet (sheetName in varchar2);
  procedure setCurrentCell (cellRef in varchar2);
  procedure setFormulaType (fmlaType in pls_integer);
  
  function getParseTree (input in varchar2, refStyle in pls_integer default REF_A1) return parseTree_t pipelined;
  
  function parse (p_expr in varchar2, p_type in pls_integer default null, p_cellRef in varchar2 default null, p_refStyle in pls_integer default null) return varchar2;
  function parseBinary (p_expr in varchar2, p_type in pls_integer default null, p_cellRef in varchar2 default null, p_refStyle in pls_integer default null) return raw;
  procedure setContext (p_sheets in ExcelTypes.CT_Sheets, p_names in ExcelTypes.CT_DefinedNames);
  function getFormula return varchar2;
  function getExternals return ExcelTypes.CT_Externals;
  function getNames return ExcelTypes.CT_DefinedNames;
  function getPtgExp (p_cellRef in varchar2) return raw;

end ExcelFmla;
/
