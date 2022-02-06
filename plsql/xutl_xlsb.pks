create or replace package xutl_xlsb is
/* ======================================================================================

  MIT License

  Copyright (c) 2018-2021 Marc Bleron

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
    Marc Bleron       2018-04-02     Creation
    Marc Bleron       2018-08-23     Bug fix : no row returned if fetch_size (p_nrows)
                                     is less than first row index
    Marc Bleron       2018-09-02     Multi-sheet support
    Marc Bleron       2020-03-01     Added cellNote attribute to ExcelTableCell
    Marc Bleron       2021-04-05     Added generation routines for ExcelGen
    Marc Bleron       2021-09-04     Added fWrap attribute
====================================================================================== */
  
  type SheetEntry_T is record (name varchar2(31 char), relId varchar2(255 char));
  type SheetEntries_T is table of SheetEntry_T;
  type SupportingLink_T is record (externalLink pls_integer, firstSheet pls_integer, lastSheet pls_integer);
  type SupportingLinks_T is table of SupportingLink_T index by pls_integer;
  
  type Stream_T is record (
    content    blob
  , sz         integer
  , offset     integer
  , rt         pls_integer
  , hsize      pls_integer
  , rsize      binary_integer
  , rstart     integer
  , available  pls_integer
  -- cache
  , buf        raw(32767)
  , buf_sz     pls_integer
  );
  
  function new_stream return Stream_T;
  procedure flush_stream (stream  in out nocopy Stream_T);
   
  procedure set_debug (p_mode in boolean);
  
  function add_SupportingLink (
    links         in out nocopy SupportingLinks_T
  , externalLink  in pls_integer
  , firstSheet    in pls_integer
  , lastSheet     in pls_integer default null
  )
  return pls_integer;
  
  procedure put_simple_record (
    stream   in out nocopy stream_t
  , recnum   in pls_integer
  , content  in raw default null
  );
  
  procedure put_RowHdr (
    stream    in out nocopy stream_t
  , rowIndex  in pls_integer
  );
  
  procedure put_BeginSst (
    stream     in out nocopy stream_t
  , cstTotal   in pls_integer
  , cstUnique  in pls_integer
  );

  procedure put_CellIsst (
    stream    in out nocopy stream_t
  , colIndex  in pls_integer
  , styleRef  in pls_integer default 0
  , isst      in pls_integer
  );
  
  procedure put_CellSt (
    stream    in out nocopy stream_t
  , colIndex  in pls_integer
  , styleRef  in pls_integer default 0
  , strValue  in varchar2 default null
  , lobValue  in varchar2 default null
  );
  
  procedure put_CellNumber (
    stream    in out nocopy stream_t
  , colIndex  in pls_integer
  , styleRef  in pls_integer default 0
  , num       in number
  );

  procedure put_SSTItem (
    stream  in out nocopy stream_t
  , str     in varchar2
  );
  
  procedure put_ExternSheet (
    stream  in out nocopy stream_t
  , links   in SupportingLinks_T
  );
  
  procedure put_defaultBookViews (
    stream  in out nocopy stream_t
  );
  
  procedure put_BundleSh (
    stream     in out nocopy stream_t
  , sheetId    in pls_integer
  , relId      in varchar2
  , sheetName  in varchar2
  );
  
  procedure put_BeginList (
    stream       in out nocopy stream_t
  , tableId      in pls_integer
  , name         in varchar2
  , displayName  in varchar2
  , showHeader   in boolean
  , firstRow     in pls_integer
  , firstCol     in pls_integer
  , lastRow      in pls_integer
  , lastCol      in pls_integer    
  );
  
  procedure put_BeginListCol (
    stream     in out nocopy stream_t
  , fieldId    in pls_integer
  , fieldName  in varchar2    
  );
  
  procedure put_BeginAFilter (
    stream    in out nocopy stream_t
  , firstRow  in pls_integer
  , firstCol  in pls_integer
  , lastRow   in pls_integer
  , lastCol   in pls_integer 
  );

  procedure put_ListPart (
    stream  in out nocopy stream_t
  , relId   in varchar2
  );
  
  procedure put_TableStyleClient (
    stream          in out nocopy stream_t
  , tableStyleName  in varchar2 
  );

  procedure put_BuiltInStyle (
    stream     in out nocopy stream_t
  , builtInId  in pls_integer
  , styleName  in varchar2
  , xfId       in pls_integer
  );

  procedure put_FilterDatabase (
    stream         in out nocopy stream_t
  , bundleShIndex  in pls_integer
  , xti            in pls_integer
  , firstRow       in pls_integer
  , firstCol       in pls_integer
  , lastRow        in pls_integer
  , lastCol        in pls_integer
  );
  
  procedure put_CalcProp (
    stream  in out nocopy stream_t
  , calcId  in pls_integer
  );

  procedure put_WsProp (
    stream    in out nocopy stream_t
  , tabColor  in varchar2
  );
  
  procedure put_BeginWsView (
    stream  in out nocopy stream_t 
  );
  
  procedure put_FrozenPane (
    stream   in out nocopy stream_t
  , numRows  in pls_integer
  , numCols  in pls_integer
  , topRow   in pls_integer  -- first row of the lower right pane
  , leftCol  in pls_integer  -- first column of the lower right pane
  );
  
  procedure put_NumFmt (
    stream  in out nocopy stream_t
  , id      in pls_integer
  , fmt     in varchar2
  );
  
  procedure put_Font (
    stream  in out nocopy stream_t
  , font    in ExcelTypes.CT_Font
  );
  
  procedure put_PatternFill (
    stream       in out nocopy stream_t
  , patternfill  in ExcelTypes.CT_PatternFill
  );
  
  procedure put_Border (
    stream  in out nocopy stream_t
  , border  in ExcelTypes.CT_Border
  );

  procedure put_XF (
    stream      in out nocopy stream_t
  , xfId        in pls_integer default null
  , numFmtId    in pls_integer default 0
  , fontId      in pls_integer default 0
  , fillId      in pls_integer default 0
  , borderId    in pls_integer default 0
  , hAlignment  in varchar2 default null
  , vAlignment  in varchar2 default null
  , wrapText    in boolean default false
  );
    
  function new_context (
    p_sst_part  in blob
  , p_cols      in varchar2 default null
  , p_firstRow  in pls_integer default null
  , p_lastRow   in pls_integer default null    
  )
  return pls_integer;

  procedure add_sheet (
    p_ctx_id    in pls_integer
  , p_content   in blob
  , p_comments  in blob
  );
  
  function get_sheetEntries (
    p_workbook  in blob
  )
  return SheetEntries_T;
 
  procedure free_context (
    p_ctx_id  in pls_integer 
  );

  function iterate_context (
    p_ctx_id  in pls_integer
  , p_nrows   in pls_integer
  )
  return ExcelTableCellList;
  
  procedure read_all (file in blob);

end xutl_xlsb;
/
