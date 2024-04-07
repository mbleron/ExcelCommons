create or replace package ExcelFmla is
/* ======================================================================================

  MIT License

  Copyright (c) 2023 Marc Bleron

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
    Marc Bleron       2020-04-01     Creation
====================================================================================== */

  EXC_FMLATYPE_CELL     constant pls_integer := 0;   -- Simple cell formula, also used in change tracking.
  EXC_FMLATYPE_NAME     constant pls_integer := 5;   -- Defined name.
  
  type parseTreeNode_t is record (
    id integer
  , nodeType integer
  , nodeTypeName varchar2(128)
  , parentId integer
  , nodeLevel integer
  , token varchar2(32767)
  , nodeClass integer
  , nodeStatus integer
  );
  
  type parseTree_t is table of parseTreeNode_t;

  subtype sheet_name is varchar2(31 char);
  type sheetList_t is table of sheet_name;

  function base26Decode (input in varchar2) return pls_integer result_cache;
  function base26Encode (input in pls_integer) return varchar2 result_cache;
  function isValidCellReference (input in varchar2) return boolean;
  
  --procedure tokenize (input in varchar2);
  function parseAll (input in varchar2) return pls_integer;
  procedure putName (value in varchar2, sheetName in varchar2 default null, idx in pls_integer default null);
  procedure putSheet (name in varchar2, idx in pls_integer default null);
  procedure setCurrentSheet (sheetName in varchar2);
  procedure setCurrentCell (cellRef in varchar2);
  procedure setFormulaType (fmlaType in pls_integer);
  
  function getParseTree (input in varchar2) return parseTree_t pipelined;
  function getFunctionalTree (input in varchar2) return parseTree_t pipelined;
  
  procedure test_eval_isect (r1 in varchar2, r2 in varchar2);
  procedure test_eval_range (list in apex_t_varchar2);
  
  function parse (p_expr in varchar2) return varchar2;
  function parseBinary (p_expr in varchar2, p_type in pls_integer default null, p_cellRef in varchar2 default null) return raw;
  procedure setContext (p_sheets in ExcelTypes.CT_Sheets, p_names in ExcelTypes.CT_DefinedNames);
  function getFormula return varchar2;
  function getExternals return ExcelTypes.CT_Externals;
  function getNames return ExcelTypes.CT_DefinedNames;

end ExcelFmla;
/
