create or replace package body xutl_xlsb is
/* ======================================================================================

  MIT License

  Copyright (c) 2018-2025 Marc Bleron

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

==========================================================================================
    Change history :
    Marc Bleron       2018-04-02     Creation
    Marc Bleron       2018-08-23     Bug fix : no row returned if fetch_size (p_nrows)
                                     is less than first row index
    Marc Bleron       2018-09-02     Multi-sheet support
    Marc Bleron       2020-03-01     Added cellNote attribute to ExcelTableCell
    Marc Bleron       2021-04-05     Added generation routines for ExcelGen
    Marc Bleron       2021-09-04     Added fWrap attribute
    Marc Bleron       2022-02-15     Added ColInfo record
    Marc Bleron       2022-08-06     Bug fix: BrtBeginWsView.fSelected set to 1 by default
                                     Added CodeInfo.ixfe
                                     Added RowHdr.ixfe, miyRw
                                     Added BrtBeginWsView.fDspGrid, fDspRwCol
                                     Added BrtTableStyleClient flags A, B, C, D
                                     Added BrtMergeCell
                                     Added BrtWsFmtInfo
    Marc Bleron       2022-11-04     Added gradientFill
    Marc Bleron       2023-02-15     Added font vertical alignment (super/sub-script)
                                     and rich text support
    Marc Bleron       2023-05-03     Added date style detection
    Marc Bleron       2024-02-23     Added font strikethrough, text rotation, indent
    Marc Bleron       2024-05-01     Added sheet state, formula support
    Marc Bleron       2024-07-03     Themed color handling in make_BrtColor
    Marc Bleron       2024-08-16     Data validation
    Marc Bleron       2024-09-06     Conditional formatting
    Marc Bleron       2025-01-26     Sheet indices update
    Marc Bleron       2025-02-14     Image support
    Marc Bleron       2025-05-08     Sheet background
========================================================================================== */

  -- Binary Record Types
  BRT_ROWHDR            constant pls_integer := 0; 
  BRT_CELLBLANK         constant pls_integer := 1;
  BRT_CELLRK            constant pls_integer := 2;
  BRT_CELLERROR         constant pls_integer := 3;
  BRT_CELLBOOL          constant pls_integer := 4;
  BRT_CELLREAL          constant pls_integer := 5;
  BRT_CELLST            constant pls_integer := 6; 
  BRT_CELLISST          constant pls_integer := 7;
  BRT_FMLASTRING        constant pls_integer := 8;
  BRT_FMLANUM           constant pls_integer := 9;
  BRT_FMLABOOL          constant pls_integer := 10;
  BRT_FMLAERROR         constant pls_integer := 11;
  BRT_SSTITEM           constant pls_integer := 19;
  BRT_FRTBEGIN          constant pls_integer := 35;
  BRT_FRTEND            constant pls_integer := 36;
  BRT_NAME              constant pls_integer := 39;
  BRT_FONT              constant pls_integer := 43;
  BRT_FMT               constant pls_integer := 44;
  BRT_FILL              constant pls_integer := 45;
  BRT_BORDER            constant pls_integer := 46;
  BRT_XF                constant pls_integer := 47;
  BRT_STYLE             constant pls_integer := 48;
  BRT_VALUEMETA         constant pls_integer := 50;
  BRT_MDB               constant pls_integer := 51;
  BRT_BEGINFMD          constant pls_integer := 52;
  BRT_ENDFMD            constant pls_integer := 53;
  BRT_COLINFO           constant pls_integer := 60;
  BRT_DVAL              constant pls_integer := 64;
  BRT_BEGINSHEET        constant pls_integer := 129;
  BRT_BEGINBOOKVIEWS    constant pls_integer := 135;
  BRT_ENDBOOKVIEWS      constant pls_integer := 136;
  BRT_BEGINWSVIEW       constant pls_integer := 137;
  BRT_ENDBUNDLESHS      constant pls_integer := 144;
  BRT_BEGINSHEETDATA    constant pls_integer := 145;
  BRT_ENDSHEETDATA      constant pls_integer := 146;
  BRT_WSPROP            constant pls_integer := 147;
  BRT_PANE              constant pls_integer := 151;
  BRT_BUNDLESH          constant pls_integer := 156;
  BRT_CALCPROP          constant pls_integer := 157;
  BRT_BOOKVIEW          constant pls_integer := 158;
  BRT_BEGINSST          constant pls_integer := 159;
  BRT_BEGINAFILTER      constant pls_integer := 161;
  BRT_MERGECELL         constant pls_integer := 176;
  BRT_BEGINSTYLESHEET   constant pls_integer := 278;
  BRT_ENDSTYLESHEET     constant pls_integer := 279;
  BRT_BEGINMETADATA     constant pls_integer := 332;
  BRT_ENDMETADATA       constant pls_integer := 333;
  BRT_BEGINESMDTINFO    constant pls_integer := 334;
  BRT_MDTINFO           constant pls_integer := 335;
  BRT_ENDESMDTINFO      constant pls_integer := 336;
  BRT_BEGINESMDB        constant pls_integer := 337;
  BRT_ENDESMDB          constant pls_integer := 338;
  BRT_BEGINESFMD        constant pls_integer := 339;
  BRT_ENDESFMD          constant pls_integer := 340;
  BRT_BEGINLIST         constant pls_integer := 343;
  BRT_BEGINLISTCOL      constant pls_integer := 347;
  BRT_BEGINEXTERNALS    constant pls_integer := 353;
  BRT_ENDEXTERNALS      constant pls_integer := 354;
  BRT_SUPSELF           constant pls_integer := 357;
  BRT_BEGINCOLINFOS     constant pls_integer := 390;
  BRT_ENDCOLINFOS       constant pls_integer := 391;
  BRT_EXTERNSHEET       constant pls_integer := 362;
  BRT_ARRFMLA           constant pls_integer := 426;
  BRT_SHRFMLA           constant pls_integer := 427;
  BRT_BEGINCONDFORMAT   constant pls_integer := 461;
  BRT_ENDCONDFORMAT     constant pls_integer := 462;
  BRT_BEGINCFRULE       constant pls_integer := 463;
  BRT_ENDCFRULE         constant pls_integer := 464;
  BRT_BEGINICONSET      constant pls_integer := 465;
  BRT_ENDICONSET        constant pls_integer := 466;
  BRT_BEGINDATABAR      constant pls_integer := 467;
  BRT_ENDDATABAR        constant pls_integer := 468;
  BRT_BEGINCOLORSCALE   constant pls_integer := 469;
  BRT_ENDCOLORSCALE     constant pls_integer := 470;
  BRT_CFVO              constant pls_integer := 471;
  BRT_WSFMTINFO         constant pls_integer := 485;
  BRT_DXF               constant pls_integer := 507;
  BRT_TABLESTYLECLIENT  constant pls_integer := 513;
  BRT_DRAWING           constant pls_integer := 550;
  BRT_BKHIM             constant pls_integer := 562;
  BRT_COLOR             constant pls_integer := 564;
  BRT_BEGINDVALS        constant pls_integer := 573;
  BRT_ENDDVALS          constant pls_integer := 574;
  BRT_BEGINFMTS         constant pls_integer := 615;
  BRT_ENDFMTS           constant pls_integer := 616;
  BRT_BEGINCELLXFS      constant pls_integer := 617;
  BRT_ENDCELLXFS        constant pls_integer := 618;
  BRT_BEGINCOMMENTS     constant pls_integer := 628;
  BRT_ENDCOMMENTLIST    constant pls_integer := 634;
  BRT_BEGINCOMMENT      constant pls_integer := 635;
  BRT_COMMENTTEXT       constant pls_integer := 637;
  BRT_LISTPART          constant pls_integer := 661;
  BRT_BEGINRICHVALUEBLOCK  constant pls_integer := 5002;
  BRT_ENDRICHVALUEBLOCK    constant pls_integer := 5003;
  
  -- Error Types
  FT_ERR_NULL         constant raw(1) := '00';
  FT_ERR_DIV_ZERO     constant raw(1) := '07';
  FT_ERR_VALUE        constant raw(1) := '0F';
  FT_ERR_REF          constant raw(1) := '17';
  FT_ERR_NAME         constant raw(1) := '1D';
  FT_ERR_NUM          constant raw(1) := '24';
  FT_ERR_NA           constant raw(1) := '2A';
  FT_ERR_GETDATA      constant raw(1) := '2B';
  
  -- Boolean Values
  BOOL_FALSE          constant raw(1) := '00';
  BOOL_TRUE           constant raw(1) := '01';
  
  -- String Types
  ST_SIMPLE           constant pls_integer := 0;
  ST_RICHSTR          constant pls_integer := 1;
  
  ERR_EXPECTED_REC    constant varchar2(100) := 'Error at position %d, expecting a [%s] record';
  
  type RecordTypeLabelMap_T is table of varchar2(128) index by pls_integer;
  recordTypeLabelMap  RecordTypeLabelMap_T;
  
  type RecordTypeByteMap_T is table of raw(2) index by pls_integer;
  recordTypeByteMap  RecordTypeByteMap_T;
  
  type BitMaskTable_T is varray(8) of raw(1);
  BITMASKTABLE    constant BitMaskTable_T := BitMaskTable_T('01','02','04','08','10','20','40','80');

  type ColumnMap_T is table of varchar2(2) index by pls_integer;
  
  type Range_T is record (
    firstRow  pls_integer
  , lastRow   pls_integer
  , colMap    ColumnMap_T
  );

  type Record_T is record (
    rt           pls_integer
  , sz           integer
  , content_raw  raw(32767)
  , is_lob       boolean := false
  , content      blob
  );
  
  type recordArray_t is table of record_t;
  
  type String_T is record (
    is_lob    boolean := false
  , strValue  varchar2(32767)
  , lobValue  clob  
  );
  
  type XLRichString_T is record (
    cch       binary_integer
  , fExtStr   boolean
  , fRichStr  boolean
  , byte_len  pls_integer := 0
  , content   String_T
  );
  
  type String_Array_T is table of String_T;
  
  type SST_T is record (
    cstTotal   binary_integer
  , cstUnique  binary_integer
  , strings    String_Array_T
  );
  
  type BrtBundleSh_T is record (
    hsState   raw(4)
  , iTabID    raw(4)
  , strRelID  varchar2(255 char)
  , strName   varchar2(31 char)
  );
  
  type Comment_T is record (
    rw    pls_integer
  , col   pls_integer
  , text  varchar2(32767)
  );
  
  type RK_T is record (
    fX100     boolean
  , fInt      boolean
  , RkNumber  raw(4)
  );
  
  type SheetList_T is table of blob;
  
  -- comments
  type CommentMap_T is table of varchar2(32767) index by varchar2(10); -- cell comment indexed by cellref
  type Comments_T is table of CommentMap_T index by pls_integer; -- comment map indexed by sheet index
  
  type dateXfMap_t is table of ExcelTypes.CT_NumFmt index by pls_integer;
  
  type sheetMap_t is table of pls_integer index by varchar2(128); -- sheet indices indexed by sheet names
  type shrFmla_t is record (ptgExp raw(32767), pos integer);
  type shrFmlaMap_t is table of shrFmla_t index by pls_integer; -- shared formula info indexed by shared index (si) 
  
  type Context_T is record (
    stream      Stream_T
  , sst         SST_T
  , rng         Range_T
  , curr_rw     pls_integer
  , done        boolean
  , sheetList   SheetList_T
  , curr_sheet  pls_integer
  , comments    Comments_T
  , dtStyles    dateXfMap_t
  );
  
  type Context_cache_T is table of Context_T index by pls_integer;
  ctx_cache  Context_cache_T;
  
  sheetMap    sheetMap_t;
  shrFmlaMap  shrFmlaMap_t;
  
  debug_mode       boolean := false;
  MAX_STRING_SIZE  pls_integer;
  
  procedure loadRecordTypeLabels is
    rc    sys_refcursor;
    rt    pls_integer;
    name  varchar2(128);
  begin
    open rc for 'select rt, name from xlsb_record_types';
    loop
      fetch rc into rt, name;
      exit when rc%notfound;
      recordTypeLabelMap(rt) := name;
    end loop;
    close rc;
  end;

  function raw2int (r in raw) return binary_integer
  is
  begin
    return utl_raw.cast_to_binary_integer(r, utl_raw.little_endian);
  end;

  function int2raw (int32 in binary_integer, sz in pls_integer default null) return raw
    result_cache
  is
    r raw(4) := utl_raw.cast_from_binary_integer(int32, utl_raw.little_endian);
  begin
    return case when sz is not null then utl_raw.substr(r, 1, sz) else r end;
  end;
  
  function bitVector (
    b0  in pls_integer default 0
  , b1  in pls_integer default 0
  , b2  in pls_integer default 0
  , b3  in pls_integer default 0
  , b4  in pls_integer default 0
  , b5  in pls_integer default 0
  , b6  in pls_integer default 0
  , b7  in pls_integer default 0
  )
  return raw
  is
  begin
    return int2raw(b0 + 2*b1 + 4*b2 + 8*b3 + 16*b4 + 32*b5 + 64*b6 + 128*b7, 1);
  end;
  
  function make_rt (recnum in pls_integer) 
  return raw
  is
    rt  raw(2);
  begin
    return recordTypeByteMap(recnum);
  exception
    when no_data_found then
      rt := case 
              when recnum >= 128 then utl_raw.concat(
                                        int2raw(mod(recnum,128)+128, 1)
                                      , int2raw(trunc(recnum/128), 1) )
              else int2raw(recnum, 1)
            end;
      recordTypeByteMap(recnum) := rt;
      return rt;
  end;
  
  function make_rsize (sz in pls_integer) 
  return raw 
  result_cache
  is
    byte1  raw(1);
    byte2  raw(1);
    byte3  raw(1);
    byte4  raw(1);
    rem    pls_integer := sz;    
  begin
    if rem < 128 then
      byte1 := int2raw(rem,1);
    else
      byte1 := int2raw(mod(rem,128)+128,1);
      rem := trunc(rem/128);
      if rem < 128 then
        byte2 := int2raw(rem,1);
      else
        byte2 := int2raw(mod(rem,128)+128,1);
        rem := trunc(rem/128);
        if rem < 128 then
          byte3 := int2raw(rem,1);
        else
          byte3 := int2raw(mod(rem,128)+128,1);
          byte4 := int2raw(trunc(rem/128),1);
        end if;
      end if;
    end if;
    return utl_raw.concat(byte1, byte2, byte3, byte4);
  end;
  
  function make_BrtColor (
    colorCode  in varchar2
  )
  return raw 
  is
    rColorCode  raw(4);
  begin
    if colorCode is not null then
      if colorCode like 'theme:%' then
        return utl_raw.concat( bitVector(0     -- fValidRGB = 0
                                       , 1, 1  -- xColorType = 3
                                       )
                             , int2raw(to_number(regexp_substr(colorCode,'\d+$')), 1)   -- index
                             , '0000'         -- nTintAndShade
                             , '00000000'     -- bRed, bGreen, bBlue, bAlpha
                             );        
      else
        rColorCode := hextoraw(colorCode);
        return utl_raw.concat( bitVector(1     -- fValidRGB = 1
                                       , 0, 1  -- xColorType = 2
                                       )
                             , '00'    -- index (ignored when xColorType = 2)
                             , '0000'  -- nTintAndShade (0 = no change)
                             , utl_raw.substr(rColorCode, 2)     -- bRed, bGreen, bBlue
                             , utl_raw.substr(rColorCode, 1, 1)  -- bAlpha
                             );
      end if;
    else
      return utl_raw.concat( '00'        -- fValidRGB = 0, xColorType = 0
                           , '00'        -- index
                           , '0000'      -- nTintAndShade
                           , '00000000'  -- bRed, bGreen, bBlue, bAlpha
                           );
    end if;
  end;

  function get_max_string_size 
  return pls_integer 
  is
    l_result  pls_integer;
  begin
    select lengthb(rpad('x',32767,'x')) 
    into l_result
    from dual;
    return l_result;
  end;

  procedure init 
  is
  begin
    MAX_STRING_SIZE := get_max_string_size();
    recordTypeLabelMap(BRT_BEGINSST) := 'BrtBeginSst';
    recordTypeLabelMap(BRT_BEGINCOMMENTS) := 'BrtBeginComments';
    recordTypeLabelMap(BRT_COMMENTTEXT) := 'BrtCommentText';
    recordTypeLabelMap(BRT_SHRFMLA) := 'BrtShrFmla';
  end;

  procedure set_debug (p_mode in boolean)
  is
  begin
    debug_mode := p_mode;  
  end;
  
  procedure debug (message in varchar2)
  is
  begin
    if debug_mode then
      dbms_output.put_line(message);
    end if;
  end;

  procedure trace_lob (message in varchar2)
  is
    type lob_stats_rec is record (
      cache_lobs     number
    , nocache_lobs   number
    , abstract_lobs  number
    );
    lob_stats  lob_stats_rec;
  begin
    select cache_lobs, nocache_lobs, abstract_lobs
    into lob_stats
    from v$temporary_lobs
    where sid = sys_context('userenv','sid');
    debug('LOB Stats for step : '||message);
    debug('cache_lobs = '||lob_stats.cache_lobs);
    debug('nocache_lobs = '||lob_stats.nocache_lobs);
    debug('abstract_lobs = '||lob_stats.abstract_lobs);
    debug('----------------------------------------------');
  end;

  procedure error (
    errcode  in pls_integer
  , message  in varchar2
  , arg1     in varchar2 default null
  , arg2     in varchar2 default null
  , arg3     in varchar2 default null
  ) 
  is
  begin
    raise_application_error(errcode, utl_lms.format_message(message, arg1, arg2, arg3));
  end;

  procedure expect (
    stream  in Stream_T
  , rt      in pls_integer
  )
  is
  begin
    if stream.rt != rt then
      error(-20731, ERR_EXPECTED_REC, stream.rstart - stream.hsize, recordTypeLabelMap(rt));
    end if;    
  end;
  
  function is_bit_set (
    byteVal  in raw
  , bitNum   in pls_integer
  )
  return boolean
  is
    bitmask  raw(1) := BITMASKTABLE(bitNum+1);
  begin
    return ( utl_raw.bit_and(byteVal, bitmask) = bitmask );
  end;

  function read_bytes (
    stream  in out nocopy Stream_T
  , amount  in pls_integer
  )
  return raw
  is
    bytes  raw(32767);
  begin
    bytes := dbms_lob.substr(stream.content, amount, stream.offset);
    stream.offset := stream.offset + amount;
    stream.available := stream.available - amount;
    return bytes;
  end;

  function read_int8 (
    stream  in out nocopy Stream_T
  )
  return binary_integer
  is
  begin
    return utl_raw.cast_to_binary_integer(read_bytes(stream, 1), utl_raw.little_endian);
  end;

  function read_int16 (
    stream  in out nocopy Stream_T
  )
  return binary_integer
  is
  begin
    return utl_raw.cast_to_binary_integer(read_bytes(stream, 2), utl_raw.little_endian);
  end;

  function read_int32 (
    stream  in out nocopy Stream_T
  )
  return binary_integer
  is
  begin
    return utl_raw.cast_to_binary_integer(read_bytes(stream, 4), utl_raw.little_endian);
  end;
  
  function new_stream
  return Stream_T
  is
    stream  Stream_T;
  begin
    dbms_lob.createtemporary(stream.content, true);
    stream.sz := 0;
    stream.buf_sz := 0;
    return stream;
  end;
  
  procedure flush_stream (
    stream  in out nocopy Stream_T 
  )
  is
  begin
    if stream.buf_sz != 0 then
      dbms_lob.writeappend(stream.content, stream.buf_sz, stream.buf);
      stream.buf := null;
      stream.buf_sz := 0;
    end if;
  end;

  function open_stream (
    wbFile  in blob
  )
  return Stream_T
  is
    stream  Stream_T;
  begin
    stream.content := wbFile;
    stream.sz := dbms_lob.getlength(stream.content);
    stream.offset := 0;
    stream.rsize := 0;
    stream.rstart := 1;
    return stream;
  end;
  
  procedure close_stream (
    stream  in out nocopy Stream_T 
  )
  is
  begin
    dbms_lob.freetemporary(stream.content);
  end;

  procedure next_record (
    stream in out nocopy Stream_T
  ) 
  is
    int8  pls_integer;
    rstart  integer;
  begin
    stream.offset := stream.rstart + stream.rsize;
    rstart := stream.offset;
    
    -- record type
    int8 := read_int8(stream);
    if int8 < 128 then
      stream.rt := int8;
    else
      stream.rt := bitand(int8,127);
      int8 := read_int8(stream);
      stream.rt := stream.rt + bitand(int8,127) * 128;
    end if;
    
    -- record size
    int8 := read_int8(stream); -- byte 1 
    stream.rsize := bitand(int8,127); -- lowest 7 bits
    if int8 >= 128 then  
      int8 := read_int8(stream); -- byte 2
      stream.rsize := stream.rsize + bitand(int8,127) * 128;
      if int8 >= 128 then
        int8 := read_int8(stream); -- byte 3
        stream.rsize := stream.rsize + bitand(int8,127) * 16384;
        if int8 >= 128 then
          int8 := read_int8(stream); -- byte 4
          stream.rsize := stream.rsize + bitand(int8,127) * 2097152;
        end if;
      end if;
    end if;
    
    stream.hsize := stream.offset - rstart;
    stream.available := stream.rsize;
    -- current record data start
    stream.rstart := stream.offset;
    --debug('RECORD INFO ['||to_char(stream.rstart,'FM0XXXXXXX')||']['||lpad(stream.rsize,6)||'] '||stream.rt);
  end;

  procedure seek_first (
    stream       in out nocopy Stream_T
  , record_type  in pls_integer  
  )
  is
  begin
    next_record(stream);
    while stream.offset < stream.sz and stream.rt != record_type loop
      next_record(stream);
    end loop;
  end;

  procedure seek (
    stream  in out nocopy Stream_T
  , pos     in integer
  )
  is
  begin
    stream.rstart := pos;
    stream.rsize := 0;
  end;
  
  procedure skip (
    stream  in out nocopy Stream_T
  , amount  in integer
  )
  is
  begin
    stream.offset := stream.offset + amount;
    stream.available := stream.available - amount;
  end;

  -- write some bytes to a raw buffer
  procedure put (
    buf    in out nocopy raw
  , bytes  in raw
  )
  is
  begin
    if buf is null then
      buf := bytes;
    else
      buf := utl_raw.overlay(bytes, buf, utl_raw.length(buf) + 1);
    end if;
  end;

  function new_record (
    rec_type     in pls_integer
  , content   in raw default null
  )
  return Record_T
  is
    rec  Record_T;
  begin
    rec.rt := rec_type;
    if content is not null then
      rec.content_raw := content;
      rec.sz := utl_raw.length(content);
    else
      rec.sz := 0;
    end if;
    return rec;
  end;

  -- write data to a record instance
  procedure write_record (
    rec  in out nocopy Record_T
  , buf  in raw
  )
  is
    buf_sz  pls_integer := nvl(utl_raw.length(buf), 0);
  begin
    if buf_sz != 0 then
      if rec.is_lob then
        dbms_lob.writeappend(rec.content, buf_sz, buf);
      elsif rec.sz + buf_sz <= 32767 then
        --rec.content_raw := utl_raw.concat(rec.content_raw, buf);
        if rec.sz = 0 then
          rec.content_raw := buf;
        else
          rec.content_raw := utl_raw.overlay(buf, rec.content_raw, rec.sz + 1);
        end if;
      else 
        -- switch to lob storage
        rec.is_lob := true;
        dbms_lob.createtemporary(rec.content, true);
        -- transfer existing raw content
        dbms_lob.writeappend(rec.content, rec.sz, rec.content_raw);
        -- append buffer content
        dbms_lob.writeappend(rec.content, buf_sz, buf);
      end if;
      rec.sz := rec.sz + buf_sz;
    end if;
  end;

  procedure put_record (
    stream  in out nocopy Stream_T
  , rec     in Record_T
  )
  is
    rt     raw(2) := make_rt(rec.rt);
    rsize  raw(4) := make_rsize(rec.sz);
    procedure put (bytes in raw) is
      len  pls_integer := nvl(utl_raw.length(bytes), 0);
    begin
      if len != 0 then
        if stream.buf_sz + len <= 1536 then
          if stream.buf_sz != 0 then
            -- overlay() is a little faster than concat() here
            stream.buf := utl_raw.overlay(bytes, stream.buf, stream.buf_sz + 1);
          else
            stream.buf := bytes;
          end if;
          stream.buf_sz := stream.buf_sz + len;
        else
          flush_stream(stream);
          stream.buf := bytes;
          stream.buf_sz := len;
        end if;
        stream.sz := stream.sz + len;
      end if;
    end;
  begin
    put(rt);
    put(rsize);
    if rec.is_lob then
      flush_stream(stream);
      dbms_lob.copy(stream.content, rec.content, dbms_lob.getlength(rec.content), dbms_lob.getlength(stream.content) + 1);
      stream.sz := stream.sz + dbms_lob.getlength(rec.content);
    else
      put(rec.content_raw);
    end if;
  end;
  
  procedure put_simple_record (
    stream  in out nocopy stream_t
  , recnum  in pls_integer
  , content in raw default null
  )
  is
  begin
    put_record(stream, new_record(recnum, content));
  end;
  
  function make_RowHdr (
    rowIndex       in pls_integer
  , height         in number
  , styleRef       in pls_integer
  , defaultHeight  in number default null
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_ROWHDR);
  begin
    write_record(rec, int2raw(rowIndex));            -- rw
    write_record(rec, int2raw(nvl(styleRef, 0)));    -- ixfe
    write_record(rec, int2raw(coalesce(height, defaultHeight, 15)*20, 2));  -- miyRw: 300 twips by default
    write_record(rec, '00');    -- A, B, reserved1
    write_record(rec, bitVector( b5 => case when height is not null then 1 else 0 end    -- fUnsynced
                               , b6 => case when styleRef is not null then 1 else 0 end  -- fGhostDirty 
                               ));  
    write_record(rec, '00');        -- I, reserved2
    write_record(rec, '00000000');  -- ccolspan    
    return rec;
  end;

  function make_Cell (
    recnum    in pls_integer
  , colIndex  in pls_integer
  , styleRef  in pls_integer
  )
  return Record_T
  is
    rec  Record_T := new_record(recnum);
  begin
    write_record(rec, int2raw(colIndex));  -- column
    write_record(rec, int2raw(styleRef));  -- iStyleRef, A, reserved
    return rec;
  end;
  
  function make_NumberCell (
    colIndex  in pls_integer
  , styleRef  in pls_integer
  , num       in number
  )
  return Record_T
  is
    rec       Record_T;
    numX100   binary_double;
    Xnum      raw(8);
    XnumX100  raw(8);
    mask      raw(8) := hextoraw('00000000FCFFFFFF');
    recnum    pls_integer := BRT_CELLRK;
  begin
    if num is not null then
      
      Xnum := utl_raw.cast_from_binary_double(num, utl_raw.little_endian);
      -- if Xnum fits in the 30 most significant bits
      if Xnum = utl_raw.bit_and(Xnum, mask) then
        Xnum := utl_raw.substr(Xnum, 5);
      -- if num is a 30-bit signed integer
      elsif trunc(num) = num and num between -536870912 and 536870911 then
        Xnum := utl_raw.cast_from_binary_integer(num*4 + 2, utl_raw.little_endian);  -- +2 : fInt=1 fX100=0
      else
        numX100 := num*100;
        -- if numX100 is an integer
        if trunc(numX100) = numX100 then
          XnumX100 := utl_raw.cast_from_binary_double(numX100, utl_raw.little_endian);
          if XnumX100 = utl_raw.bit_and(XnumX100, mask) then
            Xnum := utl_raw.bit_or(utl_raw.substr(XnumX100, 5), '01');  -- fInt=0 fX100=1
          elsif numX100 between -536870912 and 536870911 then 
            Xnum := utl_raw.cast_from_binary_integer(numX100*4 + 3, utl_raw.little_endian);  -- +3 : fInt=1 fX100=1
          else
            recnum := BRT_CELLREAL;
          end if;

        else
          recnum := BRT_CELLREAL;
        end if;
      end if;
      
      rec := make_Cell(recnum, colIndex, styleRef);
      write_record(rec, Xnum);
      return rec;
      
    else
      
      return make_Cell(BRT_CELLBLANK, colIndex, styleRef);
      
    end if;
  end;
  
  procedure write_XLWideString (
    rec       in out nocopy Record_T
  , strValue  in varchar2 default null
  , lobValue  in clob default null
  , nullable  in boolean default false
  )
  is
    cch     pls_integer;
    csname  varchar2(30) := 'AL16UTF16LE';
    amt     pls_integer;
    cbuf    varchar2(32764);
    offset  integer := 1;
    rem     integer;
  begin
    if nullable and strValue is null and lobValue is null then
      
      write_record(rec, 'FFFFFFFF'); -- NULL
    
    elsif lobValue is null then
      
      cch := nvl(length(strValue), 0);
      write_record(rec, int2raw(cch)); -- cchCharacters
      -- rgchData
      if cch < 16384 then
        write_record(rec, utl_i18n.string_to_raw(strValue, csname));
      else
        write_record(rec, utl_i18n.string_to_raw(substr(strValue, 1, 16383), csname));
        write_record(rec, utl_i18n.string_to_raw(substr(strValue, 16384, 32766), csname));
        if cch = 32767 then
          write_record(rec, utl_i18n.string_to_raw(substr(strValue, 32767, 1), csname));
        end if;
      end if;
      
    else

      cch := least(dbms_lob.getlength(lobValue), 32767);
      write_record(rec, int2raw(cch));  -- cchCharacters
      rem := cch;
      -- rgchData
      while rem != 0 loop
        amt := least(8191, rem);
        dbms_lob.read(lobValue, amt, offset, cbuf);
        offset := offset + amt;
        rem := rem - amt;
        write_record(rec, utl_i18n.string_to_raw(cbuf, csname));
      end loop;
      
    end if;
  end;  

  procedure write_RichStr (
    rec          in out nocopy Record_T
  , strValue     in varchar2
  , strRunArray  in StrRunArray_T
  )
  is
    isRichStr  boolean := ( strRunArray is not null );
  begin
    write_record(rec, bitVector(case when isRichStr then 1 else 0 end)); -- fRichStr, fExtStr, unused1
    write_XLWideString(rec, strValue);
    if isRichStr then
      write_record(rec, int2raw(strRunArray.count));  -- dwSizeStrRun
      for i in 1 .. strRunArray.count loop            -- rgsStrRun
        write_record(rec, int2raw(strRunArray(i).ich, 2));  -- ich
        write_record(rec, int2raw(strRunArray(i).ifnt, 2)); -- ifnt
      end loop;
    end if;
  end;
  
  function make_RfX (
    firstRow  in pls_integer
  , firstCol  in pls_integer
  , lastRow   in pls_integer
  , lastCol   in pls_integer 
  )
  return raw
  is
  begin
    return utl_raw.concat(
             int2raw(firstRow) -- rwFirst
           , int2raw(lastRow)  -- rwLast
           , int2raw(firstCol) -- colFirst
           , int2raw(lastCol)  -- colLast 
           );      
  end;
  
  function make_NumFmt (
    id   in pls_integer
  , fmt  in varchar2
  )
  return Record_T
  is
    numFmt  Record_T := new_record(BRT_FMT);
  begin
    write_record(numFmt, int2raw(id, 2));  -- ifmt
    write_XLWideString(numFmt, fmt);           -- stFmtCode
    return numFmt;
  end;
  
  function make_Font (
    font  in ExcelTypes.CT_Font
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_FONT);
  begin
    write_record(rec, int2raw(nvl(font.sz, ExcelTypes.DEFAULT_FONT_SIZE) * 20, 2));  -- dyHeight : font size in twips (1 twip = 1/20th pt)
    -- grbit
    write_record(rec, bitVector(
                        b1 => case when font.i then 1 else 0 end  -- bit1 : fItalic
                      , b3 => case when font.strike then 1 else 0 end --bit3 : fStrikeout
                      ));
    write_record(rec, '00'); -- FontFlags.unused3
    write_record(rec, case when font.b then 'BC02' else '9001' end);  -- bls
    write_record(rec, int2raw(ExcelTypes.getFontVerticalAlignmentId(nvl(font.vertAlign,'baseline')), 2));  -- sss
    write_record(rec, int2raw(ExcelTypes.getUnderlineStyleId(nvl(font.u,'none')), 1));  -- uls
    write_record(rec, '00');    -- bFamily : Not applicable
    write_record(rec, '01');    -- bCharset : DEFAULT_CHARSET
    write_record(rec, '00');    -- unused
    write_record(rec, make_BrtColor(font.color));
    write_record(rec, '00');    -- bFontScheme : None
    write_XLWideString(rec, nvl(font.name, ExcelTypes.DEFAULT_FONT_FAMILY));  -- name
  
    return rec;
  end;
  
  function make_PatternFill (
    patternfill  in ExcelTypes.CT_PatternFill
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_FILL);
  begin
    write_record(rec, int2raw(ExcelTypes.getFillPatternTypeId(patternFill.patternType)));  -- fls
    write_record(rec, make_BrtColor(patternFill.fgColor));  -- BrtColorFore
    write_record(rec, make_BrtColor(patternFill.bgColor));  -- BrtColorBack
    write_record(rec, utl_raw.concat( '00000000'          -- iGradientType
                                    , '0000000000000000'  -- xnumDegree
                                    , '0000000000000000'  -- xnumFillToLeft
                                    , '0000000000000000'  -- xnumFillToRight
                                    , '0000000000000000'  -- xnumFillToTop
                                    , '0000000000000000'  -- xnumFillToBottom
                                    , '00000000'          -- cNumStop
                                    ) );
    
    return rec;
  end;

  function make_GradientFill (
    gradientFill  in ExcelTypes.CT_GradientFill
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_FILL);
  begin
    write_record(rec, '00000028');  -- fls
    write_record(rec, '0000000000000000');  -- BrtColorFore
    write_record(rec, '0000000000000000');  -- BrtColorBack
    write_record(rec, '00000000');          -- iGradientType: linear
    write_record(rec, utl_raw.cast_from_binary_double(gradientFill.degree, utl_raw.little_endian)); -- xnumDegree
    write_record(rec, '0000000000000000');  -- xnumFillToLeft
    write_record(rec, '0000000000000000');  -- xnumFillToRight
    write_record(rec, '0000000000000000');  -- xnumFillToTop
    write_record(rec, '0000000000000000');  -- xnumFillToBottom
    write_record(rec, int2raw(gradientFill.stops.count));  -- cNumStop
    -- GradientStop array:
    for i in 1 .. gradientFill.stops.count loop
      write_record(rec, make_BrtColor(gradientFill.stops(i).color));  -- GradientStop.brtColor
      write_record(rec, utl_raw.cast_from_binary_double(gradientFill.stops(i).position, utl_raw.little_endian));  -- GradientStop.xnumPosition
    end loop;   
    return rec;
  end;
  
  function make_Border (
    border  in ExcelTypes.CT_Border
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_BORDER);
    function Blxf (borderPr in ExcelTypes.CT_BorderPr) return raw is
    begin
      return utl_raw.concat(
               int2raw(ExcelTypes.getBorderStyleId(nvl(borderPr.style,'none')), 1)  -- dg
             , '00'  -- reserved
             , make_BrtColor(borderPr.color)  -- brtColor
             );      
    end;
  begin
    write_record(rec, '00');                 -- fBdrDiagDown, fBdrDiagUp, reserved
    write_record(rec, Blxf(border.top));     -- blxfTop
    write_record(rec, Blxf(border.bottom));  -- blxfBottom
    write_record(rec, Blxf(border.left));    -- blxfLeft
    write_record(rec, Blxf(border.right));   -- blxfRight
    write_record(rec, Blxf(null));           -- blxfDiag
    return rec;
  end;

  function make_XF (
    xfId       in pls_integer default null
  , numFmtId   in pls_integer default 0
  , fontId     in pls_integer default 0
  , fillId     in pls_integer default 0
  , borderId   in pls_integer default 0
  , alignment  in ExcelTypes.CT_CellAlignment default null
  )
  return Record_T
  is
    xf  Record_T := new_record(BRT_XF);
  begin
    write_record(xf, int2raw(nvl(xfId, 65535),2));  -- ixfeParent
    write_record(xf, int2raw(numFmtId,2));          -- iFmt
    write_record(xf, int2raw(fontId,2));            -- iFont
    write_record(xf, int2raw(fillId,2));            -- iFill
    write_record(xf, int2raw(borderId,2));          -- ixBorder
    write_record(xf, case 
                     when alignment.textRotation is not null then int2raw(round(alignment.textRotation),1)
                     when alignment.verticalText then 'FF'
                     else '00'
                     end );  -- trot
    
    write_record(xf, int2raw(nvl(round(alignment.indent),0),1));  -- indent
        
    write_record(xf, 
      bitVector( b0 => ExcelTypes.getHorizontalAlignmentId(nvl(alignment.horizontal,'general'))  -- alc (3 bits)
               , b3 => ExcelTypes.getVerticalAlignmentId(nvl(alignment.vertical,'bottom'))     -- alcv (3 bits)
               , b6 => case when alignment.wrapText then 1 else 0 end  -- fWrap
               , b7 => 0  -- fJustLast
               )
    );
    
    write_record(xf, 
      bitVector(
        b0 => 0  -- fShrinkToFit
      , b1 => 0  -- fMergeCell
      , b2 => 0  -- iReadingOrder (2 bits)
      , b4 => 1  -- fLocked
      , b5 => 0  -- fHidden
      , b6 => 0  -- fSxButton
      , b7 => 0  -- f123Prefix
      )
    );
    
    -- xfGrbitAtr (6 bits)
    write_record(xf,
      bitVector(
        case when numFmtId != 0 then 1 else 0 end
      , case when fontId != 0 then 1 else 0 end
      , case when alignment.horizontal is not null or alignment.vertical is not null then 1 else 0 end
      , case when borderId != 0 then 1 else 0 end
      , case when fillId != 0 then 1 else 0 end
      , 0  -- applyProtection
      )
    );
    
    write_record(xf, '00');  -- unused
    
    return xf;
    
  end;
  
  function make_BuiltInStyle (
    builtInId  in pls_integer
  , styleName  in varchar2
  , xfId       in pls_integer 
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_STYLE);
  begin
    write_record(rec, int2raw(xfId));         -- ixf
    write_record(rec, '0100');                -- grbitObj1 : fBuiltIn=1
    write_record(rec, int2raw(builtInId,1));  -- iStyBuiltIn : Normal
    write_record(rec, 'FF');                  -- iLevel (ignored)
    write_XLWideString(rec, styleName);       -- stName
    return rec;
  end;

  -- 2.4.356 BrtDXF
  function make_DXF (
    dxf  in ExcelTypes.CT_Style
  )
  return record_t
  is
    rec     record_t := new_record(BRT_DXF);
    xfProp  record_t;
    cprops  pls_integer := 0;
    
    procedure put_xfProp is
    begin
      cprops := cprops + 1;
      write_record(rec, utl_raw.concat( int2raw(xfProp.rt, 2)     -- xfPropType
                                      , int2raw(xfProp.sz + 4, 2) -- cb
                                      , xfProp.content_raw        -- xfPropDataBlob
                                      ));
    end;
    
  begin
    write_record(rec, '0000'); -- unused + fNewBorder
    -- 2.5.164 XFProps
    write_record(rec, '0000'); -- reserved
    write_record(rec, '0000'); -- cprops placeholder
    -- xfPropArray [2.5.159 XFProp]
    
    if dxf.fill.fillType = ExcelTypes.FT_PATTERN then
      
      if dxf.fill.patternFill.patternType is not null then
        
        -- 2.5.51 FillPattern
        xfProp := new_record(0);
        write_record(xfProp, int2raw(ExcelTypes.getFillPatternTypeId(dxf.fill.patternFill.patternType), 1));
        put_xfProp;
        
        if dxf.fill.patternFill.fgColor is not null then
          xfProp := new_record(1);
          write_record(xfProp, make_BrtColor(dxf.fill.patternFill.fgColor));
          put_xfProp;
        end if;

        if dxf.fill.patternFill.bgColor is not null then
          xfProp := new_record(2);
          write_record(xfProp, make_BrtColor(dxf.fill.patternFill.bgColor));
          put_xfProp;
        end if;
        
      end if;
      
    elsif dxf.fill.fillType = ExcelTypes.FT_GRADIENT then
      
      -- 2.5.162 XFPropGradient
      xfProp := new_record(3);
      write_record(xfProp, '00000000');  -- type (0 = linear)
      write_record(xfProp, utl_raw.cast_from_binary_double(dxf.fill.gradientFill.degree, utl_raw.little_endian)); -- numDegree
      write_record(xfProp, '0000000000000000');  -- numFillToLeft
      write_record(xfProp, '0000000000000000');  -- numFillToRight
      write_record(xfProp, '0000000000000000');  -- numFillToTop
      write_record(xfProp, '0000000000000000');  -- numFillToBottom
      put_xfProp;
      
      -- 2.5.163 XFPropGradientStop
      for i in 1 .. dxf.fill.gradientFill.stops.count loop
        xfProp := new_record(4);
        write_record(xfProp, '0000'); -- unused
        write_record(xfProp, utl_raw.cast_from_binary_double(dxf.fill.gradientFill.stops(i).position, utl_raw.little_endian)); -- numPosition
        write_record(xfProp, make_BrtColor(dxf.fill.gradientFill.stops(i).color)); -- color
        put_xfProp;
      end loop;
          
    end if;
    
    if dxf.font.color is not null then
      xfProp := new_record(5, make_BrtColor(dxf.font.color));
      put_xfProp;
    end if;
    
    if dxf.border.top.style is not null then
      xfProp := new_record(6);
      write_record(xfProp, make_BrtColor(dxf.border.top.color));  -- color
      write_record(xfProp, int2raw(ExcelTypes.getBorderStyleId(dxf.border.top.style), 2)); -- dgBorder
      put_xfProp;
    end if;

    if dxf.border.bottom.style is not null then
      xfProp := new_record(7);
      write_record(xfProp, make_BrtColor(dxf.border.bottom.color));  -- color
      write_record(xfProp, int2raw(ExcelTypes.getBorderStyleId(dxf.border.bottom.style), 2)); -- dgBorder
      put_xfProp;
    end if;
    
    if dxf.border.left.style is not null then
      xfProp := new_record(8);
      write_record(xfProp, make_BrtColor(dxf.border.left.color));  -- color
      write_record(xfProp, int2raw(ExcelTypes.getBorderStyleId(dxf.border.left.style), 2)); -- dgBorder
      put_xfProp;
    end if;
    
    if dxf.border.right.style is not null then
      xfProp := new_record(9);
      write_record(xfProp, make_BrtColor(dxf.border.right.color));  -- color
      write_record(xfProp, int2raw(ExcelTypes.getBorderStyleId(dxf.border.right.style), 2)); -- dgBorder
      put_xfProp;
    end if;
    
    if dxf.alignment.horizontal is not null then
      xfProp := new_record(15, int2raw(ExcelTypes.getHorizontalAlignmentId(dxf.alignment.horizontal), 1));
      put_xfProp;
    end if;

    if dxf.alignment.vertical is not null then
      xfProp := new_record(16, int2raw(ExcelTypes.getVerticalAlignmentId(dxf.alignment.vertical), 1));
      put_xfProp;
    end if;

    if dxf.alignment.textRotation is not null or dxf.alignment.verticalText then
      xfProp := new_record(17);
      if dxf.alignment.textRotation is not null then
        write_record(xfProp, int2raw(round(dxf.alignment.textRotation), 1));
      else
        write_record(xfProp, 'FF');
      end if;  -- trot
      put_xfProp;
    end if;
    
    if dxf.alignment.indent is not null then
      xfProp := new_record(18, int2raw(round(dxf.alignment.indent), 2));
      put_xfProp;
    end if;
    
    if dxf.alignment.wrapText then
      xfProp := new_record(20, '01');
      put_xfProp;
    end if;

    if dxf.font.name is not null then
      xfProp := new_record(24);
      write_record(xfProp, int2raw(length(dxf.font.name), 2)); -- LPWideString.cchCharacters
      write_record(xfProp, utl_i18n.string_to_raw(dxf.font.name, 'AL16UTF16LE')); -- LPWideString.rgchData
      put_xfProp;
    end if;
    
    if dxf.font.b is not null then
      -- 2.5.6 Bold
      xfProp := new_record(25, case when dxf.font.b then 'BC02' else '9001' end);
      put_xfProp;
    end if;
    
    if dxf.font.u is not null then
      -- 2.5.157 Underline
      xfProp := new_record(26, int2raw(ExcelTypes.getUnderlineStyleId(dxf.font.u), 2));
      put_xfProp;
    end if;
    
    if dxf.font.vertAlign is not null then
      -- 2.5.131 Script
      xfProp := new_record(27, int2raw(ExcelTypes.getFontVerticalAlignmentId(dxf.font.vertAlign), 2));
      put_xfProp;
    end if;
    
    if dxf.font.i is not null then
      xfProp := new_record(28, case when dxf.font.i then '01' else '00' end);
      put_xfProp;
    end if;
    
    if dxf.font.strike is not null then
      xfProp := new_record(29, case when dxf.font.strike then '01' else '00' end);
      put_xfProp;
    end if;
    
    if dxf.font.sz is not null then
      xfProp := new_record(36, int2raw(dxf.font.sz * 20, 4));
      put_xfProp;
    end if;
    
    if dxf.numberFormat is not null then
      xfProp := new_record(38);
      write_record(xfProp, int2raw(length(dxf.numberFormat), 2));
      write_record(xfProp, utl_i18n.string_to_raw(dxf.numberFormat, 'AL16UTF16LE'));
      put_xfProp;
      
      xfProp := new_record(41, int2raw(dxf.numFmtId, 2));
      put_xfProp;
    end if;
    
    -- update cprops placeholder
    rec.content_raw := utl_raw.overlay(int2raw(cprops, 2), rec.content_raw, 5);
    
    return rec;
    
  end;
  
  function make_BundleSh (
    sheetId    in pls_integer
  , relId      in varchar2
  , sheetName  in varchar2
  , state      in pls_integer
  )
  return Record_T
  is
    sh  Record_T := new_record(BRT_BUNDLESH);
  begin
    write_record(sh, int2raw(state));    -- hsState
    write_record(sh, int2raw(sheetId));  -- iTabID
    write_XLWideString(sh, relId);       -- strRelID
    write_XLWideString(sh, sheetName);   -- strName
    return sh;
  end;
  
  function make_ExternSheet (
    xtiArray  in ExcelTypes.xtiArray_t
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_EXTERNSHEET);
  begin
    write_record(rec, int2raw(xtiArray.count));  -- cXti
    for i in 1 .. xtiArray.count loop
      -- 2.5.173 Xti : 
      write_record(rec, int2raw(xtiArray(i).externalLink));
      write_record(rec, int2raw(xtiArray(i).firstSheet.idx));
      write_record(rec, int2raw(xtiArray(i).lastSheet.idx));
    end loop;
    return rec;
  end;
  
  function make_CalcProp (
    calcId    in pls_integer
  , refStyle  in pls_integer
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_CALCPROP);
  begin
    write_record(rec, int2raw(calcId)); -- recalcID
    write_record(rec, int2raw(1)); -- fAutoRecalc (auto)
    write_record(rec, int2raw(100)); -- cCalcCount
    write_record(rec, utl_raw.cast_from_binary_double(.001, utl_raw.little_endian)); -- xnumDelta
    write_record(rec, int2raw(1)); -- cUserThreadCount
    write_record(rec,
      bitVector(
        0 -- A: fFullCalcOnLoad
      , refStyle -- B: fRefA1
      , 0 -- C: fIter
      , 1 -- D: fFullPrec
      , 0 -- E: fSomeUncalced 
      , 1 -- F: fSaveRecalc
      , 1 -- G: fMTREnabled
      , 0 -- H: fUserSetThreadCount
      )
    );
    write_record(rec,
      bitVector(
        0 -- I: fNoDeps
      )
    );
    return rec;
  end;
  
  function make_WsProp (
    tabColor  in varchar2
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_WSPROP);
  begin
    write_record(rec,
      bitVector(
        1     -- A: fShowAutoBreaks
      , 0, 0  -- B: reserved1
      , 1     -- C: fPublish
      , 0     -- D: fDialog
      , 0     -- E: fApplyStyles
      , 1     -- F: fRowSumsBelow
      , 1     -- G: fColSumsRight
      )
    );
       
    write_record(rec,
      bitVector(
        0  -- H: fFitToPage
      , 0  -- I: reserved2
      , 1  -- J: fShowOutlineSymbols
      , 0  -- K: reserved3
      , 0  -- L: fSyncHoriz
      , 0  -- M: fSyncVert
      , 0  -- N: fAltExprEval
      , 0  -- O: fAltFormulaEntry
      )
    );
    
    write_record(rec,
      bitVector(
        0  -- P: fFilterMode
      , 1  -- Q: fCondFmtCalc
      , 0, 0, 0, 0, 0, 0  -- reserved4
      )
    );
    
    write_record(rec, make_BrtColor(tabColor));  -- brtcolorTab
    write_record(rec, 'FFFFFFFF');  -- rwSync
    write_record(rec, 'FFFFFFFF');  -- colSync
    write_record(rec, '00000000');  -- strName
  
    return rec;
  end;
  
  function make_BeginWsView (
    dspGrid   in boolean
  , dspRwCol  in boolean
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_BEGINWSVIEW);
  begin
    write_record(rec,
      bitVector(
        0  -- fWnProt
      , 0  -- fDspFmla
      , case when not dspGrid then 0 else 1 end   -- fDspGrid
      , case when not dspRwCol then 0 else 1 end  -- fDspRwCol
      , 1  -- fDspZeros
      , 0  -- fRightToLeft
      , 0  -- fSelected
      , 1  -- fDspRuler
      )
    );
    
    write_record(rec,
      bitVector(
        1  -- fDspGuts
      , 1  -- fDefaultHdr
      , 0  -- fWhitespaceHidden
      )
    );
    
    write_record(rec, '00000000');  -- xlView : XLVNORMAL
    write_record(rec, '00000000');  -- rwTop
    write_record(rec, '00000000');  -- colLeft
    write_record(rec, '40');        -- icvHdr : icvForeground (but ignored)
    write_record(rec, '000000');    -- reserved2, reserved3
    write_record(rec, '6400');      -- wScale : value?
    write_record(rec, '0000');      -- wScaleNormal : 100%
    write_record(rec, '0000');      -- wScaleSLV
    write_record(rec, '0000');      -- wScalePLV
    write_record(rec, '00000000');  -- iWbkView
    
    return rec;
  end;
  
  function make_FrozenPane (
    numRows  in pls_integer
  , numCols  in pls_integer
  , topRow   in pls_integer  -- first row of the lower right pane
  , leftCol  in pls_integer  -- first column of the lower right pane
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_PANE);
  begin
    write_record(rec, utl_raw.cast_from_binary_double(numCols, utl_raw.little_endian));  -- xnumXSplit
    write_record(rec, utl_raw.cast_from_binary_double(numRows, utl_raw.little_endian));  -- xnumYSplit
    write_record(rec, int2raw(topRow));   -- rwTop
    write_record(rec, int2raw(leftCol));  -- colLeft
    write_record(rec, '02000000');        -- pnnAct : PNNBOTLEFT
    write_record(rec, bitVector(1, 0));   -- fFrozen, fFrozenNoSplit 
    
    return rec;
  end;
  
  function make_BeginAFilter (
    firstRow  in pls_integer
  , firstCol  in pls_integer
  , lastRow   in pls_integer
  , lastCol   in pls_integer 
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_BEGINAFILTER);
  begin
    write_record(rec, make_RfX(firstRow, firstCol, lastRow, lastCol));
    return rec;
  end;
  
  function make_ListPart (
    relId  in varchar2
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_LISTPART);
  begin
    write_XLWideString(rec, relId);
    return rec;
  end;
  
  function make_BeginList (
    tableId      in pls_integer
  , name         in varchar2
  , displayName  in varchar2
  , showHeader   in boolean
  , firstRow     in pls_integer
  , firstCol     in pls_integer
  , lastRow      in pls_integer
  , lastCol      in pls_integer    
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_BEGINLIST);
  begin
    write_record(rec, make_RfX(firstRow, firstCol, lastRow, lastCol));  -- rfxList
    write_record(rec, '00000000');        -- lt : LTRANGE
    write_record(rec, int2raw(tableId));  -- idList
    write_record(rec, int2raw(case when showHeader then 1 else 0 end));  -- crwHeader
    write_record(rec, '00000000');  -- crwTotals
    write_record(rec, bitVector(0   -- fShownTotalRow
                              , 0   -- fSingleCell
                              , 0   -- fForceInsertToBeVisible
                              , 0   -- fInsertRowInsCells
                              , 0   -- fPublished
                               ));
    write_record(rec, '000000');    -- reserved
    write_record(rec, 'FFFFFFFF');  -- nDxfHeader
    write_record(rec, 'FFFFFFFF');  -- nDxfData
    write_record(rec, 'FFFFFFFF');  -- nDxfAgg
    write_record(rec, 'FFFFFFFF');  -- nDxfBorder
    write_record(rec, 'FFFFFFFF');  -- nDxfHeaderBorder
    write_record(rec, 'FFFFFFFF');  -- nDxfAggBorder
    write_record(rec, '00000000');  -- dwConnID
    write_XLWideString(rec, name);         -- stName
    write_XLWideString(rec, displayName);  -- stDisplayName
    write_record(rec, 'FFFFFFFF');  -- stComment : NULL
    write_record(rec, 'FFFFFFFF');  -- stStyleHeader : NULL
    write_record(rec, 'FFFFFFFF');  -- stStyleData : NULL
    write_record(rec, 'FFFFFFFF');  -- stStyleAgg : NULL
    
    return rec;
  end;
  
  function make_BeginListCol (
    fieldId    in pls_integer
  , fieldName  in varchar2
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_BEGINLISTCOL);
  begin
    write_record(rec, int2raw(fieldId));  -- idField
    write_record(rec, '00000000');   -- ilta : ILTA_NONE
    write_record(rec, 'FFFFFFFF');   -- nDxfHdr
    write_record(rec, 'FFFFFFFF');   -- nDxfInsertRow
    write_record(rec, 'FFFFFFFF');   -- nDxfAgg
    write_record(rec, '00000000');   -- idqsif
    write_record(rec, 'FFFFFFFF');   -- stName : NULL
    write_XLWideString(rec, fieldName);  -- stCaption
    write_record(rec, 'FFFFFFFF');  -- stTotal : NULL
    write_record(rec, 'FFFFFFFF');  -- stStyleHeader : NULL
    write_record(rec, 'FFFFFFFF');  -- stStyleInsertRow : NULL
    write_record(rec, 'FFFFFFFF');  -- stStyleAgg : NULL
    return rec;
  end;
  
  function make_TableStyleClient (
    tableStyleName     in varchar2
  , showFirstColumn    in boolean
  , showLastColumn     in boolean
  , showRowStripes     in boolean
  , showColumnStripes  in boolean
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_TABLESTYLECLIENT);
  begin
    write_record(rec, 
      bitVector(
        case when showFirstColumn then 1 else 0 end  -- fFirstColumn
      , case when showLastColumn then 1 else 0 end   -- fLastColumn
      , case when showRowStripes then 1 else 0 end   -- fRowStripes
      , case when showColumnStripes then 1 else 0 end  -- fColumnStripes
      , 0  -- fRowHeaders
      , 0  --fColumnHeaders
      )
    );
    write_record(rec, '00');  -- reserved
    write_XLWideString(rec, tableStyleName, nullable => true); -- stStyleName
    
    return rec;
  end;

  function make_ColInfo (
    colId          in pls_integer
  , colWidth       in number
  , isCustomWidth  in boolean
  , styleRef       in pls_integer
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_COLINFO);
  begin
    write_record(rec, int2raw(colId));  -- colFirst
    write_record(rec, int2raw(colId));  -- colLast
    write_record(rec, int2raw(colWidth * 256));  -- coldx
    write_record(rec, int2raw(styleRef));  -- ixfe
    write_record(rec,
      bitVector(
        0  -- fHidden
      , case when isCustomWidth then 1 else 0 end  -- fUserSet
      , 0  -- fBestFit
      , 0  -- fPhonetic
      )
    );
    write_record(rec, '00'); -- iOutLevel, unused, fCollapsed, reserved2
        
    return rec;
  end;

  function make_MergeCell (
    rwFirst   in pls_integer
  , rwLast    in pls_integer
  , colFirst  in pls_integer
  , colLast   in pls_integer
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_MERGECELL);
  begin
    write_record(rec, int2raw(rwFirst));
    write_record(rec, int2raw(rwLast));
    write_record(rec, int2raw(colFirst));
    write_record(rec, int2raw(colLast));
    return rec;
  end;

  function make_WsFmtInfo (
    defaultRowHeight  in number
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_WSFMTINFO);
  begin
    write_record(rec, 'FFFFFFFF'); -- dxGCol
    write_record(rec, int2raw(10, 2)); -- cchDefColWidth
    write_record(rec, int2raw(nvl(defaultRowHeight, 15)*20, 2)); -- miyDefRwHeight
    write_record(rec
               , bitVector(
                   case when defaultRowHeight is not null then 1 else 0 end  -- fUnsynced
                 , 0  -- fDyZero
                 , 0  -- fExAsc
                 , 0  -- fExDesc
                 ) );
    write_record(rec, '00'); -- reserved
    write_record(rec, '0000'); -- iOutLevelRw, iOutLevelCol 
    return rec;
  end;
  
  -- 2.4.711 BrtName
  function make_Name (
    names  in out nocopy ExcelTypes.CT_DefinedNames
  , idx    in pls_integer
  )
  return record_t
  is
    nm        ExcelTypes.CT_DefinedName := names(idx);
    rec       record_t := new_record(BRT_NAME);
    newNames  ExcelTypes.CT_DefinedNames;
  begin
    write_record(rec
               , bitVector(
                   case when nm.hidden then 1 else 0 end  -- fHidden
                 , case when nm.futureFunction then 1 else 0 end  -- fFunc (implied if fFutureFunction=1)
                 , 0  -- fOB
                 , case when nm.futureFunction then 1 else 0 end  -- fProc (implied if fFutureFunction=1)
                 , 0  -- fCalcExp
                 , case when nm.builtIn then 1 else 0 end  -- fBuiltin
                 , 0  -- fgrp[0]
                 , 0  -- fgrp[1]
                 ));
    write_record(rec
               , bitVector(
                   0  -- fgrp[2]
                 , 0  -- fgrp[3]
                 , 0  -- fgrp[4]
                 , 0  -- fgrp[5]
                 , 0  -- fgrp[6]
                 , 0  -- fgrp[7]
                 , 0  -- fgrp[8]
                 , 0  -- fPublished
                 ));
    write_record(rec
               , bitVector(
                   0  -- fWorkbookParam
                 , case when nm.futureFunction then 1 else 0 end  -- fFutureFunction
                 , 0  -- reserved[0]
                 , 0  -- reserved[1]
                 , 0  -- reserved[2]
                 , 0  -- reserved[3]
                 , 0  -- reserved[4]
                 , 0  -- reserved[5]
                 ));
    write_record(rec, '00'); -- reserved[6 to 13]
    write_record(rec, '00'); -- chKey
    -- itab
    if nm.scope is not null then
      write_record(rec, int2raw(sheetMap(upper(nm.scope))));
    else
      write_record(rec, 'FFFFFFFF');
    end if;
    write_XLWideString(rec, nm.name); -- name
    
    -- set current sheet to resolve unscoped cell references in the formula,
    -- an error will be raised during parsing if an unscoped cell reference is found in a workbook-level name
    ExcelFmla.setCurrentSheet(nm.scope);
    
    write_record(rec, ExcelFmla.parseBinary(nm.formula, ExcelFmla.FMLATYPE_NAME, nm.cellRef, nm.refStyle));
    write_XLWideString(rec, nm.comment, nullable => true); -- comment
    
    -- if fProc = 1 (which is implied by fFutureFunction)
    if nm.futureFunction then
      write_record(rec, 'FFFFFFFF'); -- unusedstring1
      write_record(rec, 'FFFFFFFF'); -- description
      write_record(rec, 'FFFFFFFF'); -- helpTopic
      write_record(rec, 'FFFFFFFF'); -- unusedstring2
    end if;
    
    -- retrieving new names from formula context
    newNames := ExcelFmla.getNames;
    if newNames.count > names.count then
      for i in names.count + 1 .. newNames.count loop
        names.extend;
        names(i) := newNames(i);
      end loop;
    end if;
    
    return rec;

  end;
  
  function make_Fmla (
    colIndex  in pls_integer
  , styleRef  in pls_integer
  , expr      in varchar2
  , si        in pls_integer
  , cellRef   in varchar2
  , refStyle  in pls_integer
  )
  return Record_T
  is
    rec  record_t := make_Cell(BRT_FMLAERROR, colIndex, styleRef);
  begin
    write_record(rec, FT_ERR_NA);
    write_record(rec, bitVector(b1 => 1)); -- GrbitFmla.fAlwaysCalc
    write_record(rec, '00'); -- unused
    
    if si is null then
      write_record(rec, ExcelFmla.parseBinary(expr, ExcelFmla.FMLATYPE_CELL, cellRef, refStyle));
    else
      write_record(rec, shrFmlaMap(si).ptgExp);
    end if;
    
    return rec;
  end;

  -- 2.4.785 BrtShrFmla
  function make_ShrFmla (
    expr      in varchar2
  , cellRef   in varchar2
  , refStyle  in pls_integer
  )
  return Record_T
  is
    rec  record_t := new_record(BRT_SHRFMLA);
  begin
    write_record(rec, utl_raw.copies('00',16)); -- rfx placeholder    
    write_record(rec, ExcelFmla.parseBinary(expr, ExcelFmla.FMLATYPE_SHARED, cellRef, refStyle));
    return rec;
  end;

  -- 2.4.55 BrtBeginDVals
  function make_BeginDVals (
    cnt  in pls_integer
  )
  return record_t
  is
    rec  record_t := new_record(BRT_BEGINDVALS);
  begin
    write_record(rec, '0000'); -- A + reserved
    write_record(rec, '00000000'); -- xLeft
    write_record(rec, '00000000'); -- yTop
    write_record(rec, '00000000'); -- unused3
    write_record(rec, int2raw(cnt)); -- idvMac
    return rec;
  end;

  -- 2.4.353 BrtDVal
  function make_DVal (
    dvRule  in ExcelTypes.CT_DataValidation
  )
  return record_t
  is
    rec      record_t := new_record(BRT_DVAL);
    valType  boolean;
  begin
    write_record(rec
               , int2raw(
                   ExcelTypes.getDataValidationTypeId(dvRule.type) -- valType (4 bits)
                   + 16 * ExcelTypes.getDataValidationErrStyleId(dvRule.errorStyle) -- errStyle (3 bits) << 4
                 , 1
                 )
               );
    write_record(rec
               , bitVector(
                   case when dvRule.allowBlank then 1 else 0 end   -- fAllowBlank
                 , case when dvRule.showDropDown then 0 else 1 end -- fSuppressCombo
                 , 0,0,0,0,0,0 -- bits 0-5 of mdImeMode (0x00 = No control)
                 ));
    write_record(rec
               , utl_raw.bit_or(
                   bitVector(
                     0,0 -- bits 6-7 of mdImeMode
                   , case when dvRule.showInputMessage then 1 else 0 end -- fShowInputMsg
                   , case when dvRule.showErrorMessage then 1 else 0 end -- fShowErrorMsg
                   )
                 , int2raw(16 * ExcelTypes.getDataValidationOpId(dvRule.operator), 1) -- typOperator << 4
                 ));
    write_record(rec, '00'); -- reserved
    write_record(rec, int2raw(dvRule.sqref.ranges.count)); -- sqrfx.crfx
    for i in 1 .. dvRule.sqref.ranges.count loop
      write_record(rec, int2raw(dvRule.sqref.ranges(i).rwFirst - 1)); -- rgrfx[i].rwFirst
      write_record(rec, int2raw(dvRule.sqref.ranges(i).rwLast - 1)); -- rgrfx[i].rwLast
      write_record(rec, int2raw(dvRule.sqref.ranges(i).colFirst - 1)); -- rgrfx[i].colFirst
      write_record(rec, int2raw(dvRule.sqref.ranges(i).colLast - 1)); -- rgrfx[i].colLast
    end loop;
    
    write_XLWideString(rec, dvRule.errorTitle, nullable => true);  -- DValStrings.strErrorTitle
    write_XLWideString(rec, dvRule.error, nullable => true);       -- DValStrings.strError
    write_XLWideString(rec, dvRule.promptTitle, nullable => true); -- DValStrings.strPromptTitle
    write_XLWideString(rec, dvRule.prompt, nullable => true);      -- DValStrings.strPrompt
    
    -- 2.5.98.8 DVParsedFormula
    -- root VALUE_TYPE forced to false if this is a list-based validation
    if dvRule.type = 'list' then
      valType := false;
    end if;
    
    write_record(rec, ExcelFmla.parseBinary(dvRule.fmla1, ExcelFmla.FMLATYPE_DATAVAL, dvRule.sqref.lastRangeCellRef, dvRule.refStyle1, valType));
    write_record(rec, ExcelFmla.parseBinary(dvRule.fmla2, ExcelFmla.FMLATYPE_DATAVAL, dvRule.sqref.lastRangeCellRef, dvRule.refStyle2));
        
    return rec;
  end;
  
  -- 2.4.23 BrtBeginCFRule
  function make_BeginCFRule (
    rule  in ExcelTypes.CT_CfRule
  , pos   in pls_integer
  )
  return record_T
  is
    rec  record_t := new_record(BRT_BEGINCFRULE);
  begin
    write_record(rec, int2raw(rule.type)); -- iType
    write_record(rec, int2raw(rule.template)); -- iTemplate
    write_record(rec, int2raw(nvl(rule.dxfId, -1))); -- dxfId
    write_record(rec, int2raw(pos)); -- iPri
    write_record(rec, int2raw(nvl(rule.param, 0))); -- iParam
    write_record(rec, '0000000000000000'); -- reserved1 + reserved2
    write_record(rec, bitVector(b0 => 0 -- reserved3
                              , b1 => case when rule.stopTrue then 1 else 0 end -- fStopTrue
                              , b2 => case when rule.template in (ExcelTypes.CF_TEMP_ABOVEAVERAGE, ExcelTypes.CF_TEMP_EQUALABOVEAVERAGE) then 1 else 0 end -- fAbove
                              , b3 => case when rule.bottom then 1 else 0 end -- fBottom
                              , b4 => case when rule.percent then 1 else 0 end -- fPercent
                              ));
    write_record(rec, '00'); -- last 8 bits of reserved4
    
    write_record(rec, int2raw(case when rule.fmla1 is null then 0 else 1 end)); -- cbFmla1
    write_record(rec, int2raw(case when rule.fmla2 is null then 0 else 1 end)); -- cbFmla2
    write_record(rec, int2raw(case when rule.fmla3 is null then 0 else 1 end)); -- cbFmla3
    
    write_XLWideString(rec, rule.strParam, nullable => true); -- strParam
    
    -- rgce1
    if rule.fmla1 is not null then
      write_record(rec, ExcelFmla.parseBinary(rule.fmla1, ExcelFmla.FMLATYPE_CONDFMT, rule.sqref.boundingAreaFirstCellRef, rule.refStyle1));
    end if;

    -- rgce2
    if rule.fmla2 is not null then
      write_record(rec, ExcelFmla.parseBinary(rule.fmla2, ExcelFmla.FMLATYPE_CONDFMT, rule.sqref.boundingAreaFirstCellRef, rule.refStyle2));
    end if;
    
    -- rgce3
    if rule.fmla3 is not null then
      write_record(rec, ExcelFmla.parseBinary(rule.fmla3, ExcelFmla.FMLATYPE_CONDFMT, rule.sqref.boundingAreaFirstCellRef, rule.refStyle3));
    end if;
    
    return rec;
    
  end;
  
  -- 2.4.331 BrtCFVO
  function make_CFVO (
    cfvo    in ExcelTypes.CT_Cfvo
  , cfType  in pls_integer
  )
  return record_t
  is
    decimalSep  constant varchar2(1) := substr(ExcelFmla.getNLS('NLS_NUMERIC_CHARACTERS'), 1, 1);
    rec         record_t := new_record(BRT_CFVO);
    numParam    binary_double;
    fmla        raw(32767);
    cbFmla      raw(4);
  begin
    write_record(rec, int2raw(cfvo.type)); -- iType
    
    if cfvo.type in (ExcelTypes.CFVO_NUM, ExcelTypes.CFVO_PERCENT, ExcelTypes.CFVO_PERCENTILE) then
      -- check if this CFVO value is a valid number
      begin
        numParam := to_binary_double(replace(cfvo.value, '.', decimalSep));
      exception
        when value_error then
          null;
      end;
    end if;
    write_record(rec, utl_raw.cast_from_binary_double(nvl(numParam, 0), utl_raw.little_endian)); -- numParam
    write_record(rec, int2raw(case when cfType = ExcelTypes.CF_TYPE_MULTISTATE then 1 else 0 end)); -- fSaveGTE
    write_record(rec, int2raw(case when nvl(cfvo.gte, true) then 1 else 0 end)); -- fGTE
    
    if numParam is null then
      fmla := ExcelFmla.parseBinary(cfvo.value, ExcelFmla.FMLATYPE_CELL, p_refStyle => cfvo.refStyle);
      cbFmla := utl_raw.substr(fmla, 1, 4); -- cbFmla = formula.cce
    else
      cbFmla := '00000000';
    end if;
    
    write_record(rec, cbFmla);
    write_record(rec, fmla); -- formula
    
    return rec;
  end;
  
  -- 2.4.43 BrtBeginDatabar
  function make_BeginDatabar (
    showValue  in boolean
  )
  return record_t
  is
    rec   record_t := new_record(BRT_BEGINDATABAR);
  begin
    write_record(rec, int2raw(10,1)); -- bLenMin: 10% by default
    write_record(rec, int2raw(90,1)); -- bLenMax: 90% by default
    write_record(rec, case when showValue then '01' else '00' end); -- fShowValue
    return rec;
  end;

  -- 2.4.91 BrtBeginIconSet
  function make_BeginIconSet (
    iconSet    in pls_integer
  , hideValue  in boolean
  , reverse    in boolean
  )
  return record_t
  is
    rec   record_t := new_record(BRT_BEGINICONSET);
  begin
    write_record(rec, int2raw(iconSet)); -- iSet
    write_record(rec, bitVector(
                        b0 => 0 -- reserved1
                      , b1 => case when hideValue then 1 else 0 end -- fIcon
                      , b2 => case when reverse then 1 else 0 end -- fReverse (warning! the specs have it the wrong way around)
                      --, 0, 0, 0, 0 -- unused1-4 
                      ));
    write_record(rec, '00'); -- reserved2
    return rec;
  end;
  
  -- 2.4.34 BrtBeginConditionalFormatting
  function make_BeginCondFormat (
    cfRule  in ExcelTypes.CT_CfRule
  )
  return record_t
  is
    rec   record_t := new_record(BRT_BEGINCONDFORMAT);
  begin
    write_record(rec, int2raw(1)); -- ccf
    write_record(rec, '00000000'); -- fPivot
    -- sqrfx
    -- 2.5.156 UncheckedSqRfX
    write_record(rec, int2raw(cfRule.sqref.ranges.count)); -- crfx
    for i in 1 .. cfRule.sqref.ranges.count loop
      write_record(rec, make_RfX(
                          cfRule.sqref.ranges(i).rwFirst - 1
                        , cfRule.sqref.ranges(i).colFirst - 1
                        , cfRule.sqref.ranges(i).rwLast - 1
                        , cfRule.sqref.ranges(i).colLast - 1
                        ));
    end loop;
    return rec;
  end;

  -- 2.4.704 BrtMdtinfo
  function make_MdtInfo (
    mdtName  in varchar2 
  )
  return record_t
  is
    rec record_t := new_record(BRT_MDTINFO);
  begin
    --2.5.95 MdtFlags 
    write_record(rec
               , bitVector(
                   0  -- fGhostRw
                 , 0  -- fGhostCol
                 , 0  -- fEdit
                 , 0  -- fDelete
                 , 1  -- fCopy
                 , 1  -- fPasteAll
                 , 0  -- fPasteFmlas
                 , 1  -- fPasteValues
                 ));
    write_record(rec
               , bitVector(
                   0  -- fPasteFmts
                 , 0  -- fPasteComments
                 , 0  -- fPasteDv
                 , 0  -- fPasteBorders
                 , 0  -- fPasteColWidths
                 , 0  -- fPasteNumFmts
                 , 1  -- fMerge
                 , 1  -- fSplitFirst
                 ));                 
    write_record(rec
               , bitVector(
                   0  -- fSplitAll
                 , 1  -- fRwColShift
                 , 0  -- fClearAll
                 , 1  -- fClearFmts
                 , 0  -- fClearContents
                 , 1  -- fClearComments
                 , 1  -- fAssign
                 , 0  -- reserved1 - bit0
                 )); 
    write_record(rec
               , bitVector(
                   0  -- reserved1 - bit1
                 , 0  -- reserved1 - bit2
                 , 0  -- reserved1 - bit3
                 , 0  -- reserved2 -- /!\ specs say it MUST 0 though Excel sets it to 1
                 , 1  -- fCanCoerce
                 , 0  -- fAdjust
                 , 0  -- fCellMeta
                 , 1  -- reserved3
                 ));
    write_record(rec, int2raw(120000));     -- metadataID
    write_XLWideString(rec, mdtName); -- stName
    return rec;
  end;
  
  -- 2.4.73 BrtBeginEsfmd
  function make_BeginEsfmd (
    recCount  in pls_integer
  , mdtName   in varchar2
  )
  return record_t
  is
    rec record_t := new_record(BRT_BEGINESFMD);
  begin
    write_record(rec, int2raw(recCount)); -- cFmd
    write_XLWideString(rec, mdtName);     -- stName
    return rec;
  end;
  
  procedure put_RowHdr (
    stream         in out nocopy stream_t
  , rowIndex       in pls_integer
  , height         in number
  , styleRef       in pls_integer
  , defaultHeight  in number default null
  )
  is
  begin
    put_record(stream, make_RowHdr(rowIndex, height, styleRef, defaultHeight));
  end;

  procedure put_CellNumber (
    stream    in out nocopy stream_t
  , colIndex  in pls_integer
  , styleRef  in pls_integer default 0
  , num       in number
  )
  is
  begin
    put_record(stream, make_NumberCell(colIndex, styleRef, num));
  end;
  
  procedure put_CellIsst (
    stream    in out nocopy stream_t
  , colIndex  in pls_integer
  , styleRef  in pls_integer default 0
  , isst      in pls_integer
  )
  is
    rec  Record_T := make_Cell(BRT_CELLISST, colIndex, styleRef);
  begin
    write_record(rec, int2raw(isst));
    put_record(stream, rec);
  end;
  
  procedure put_CellSt (
    stream    in out nocopy stream_t
  , colIndex  in pls_integer
  , styleRef  in pls_integer default 0
  , strValue  in varchar2 default null
  , lobValue  in varchar2 default null
  )
  is
    rec  Record_T := make_Cell(BRT_CELLST, colIndex, styleRef);
  begin
    write_XLWideString(rec, strValue, lobValue);
    put_record(stream, rec);
  end;
  
  procedure put_BeginSst (
    stream     in out nocopy stream_t
  , cstTotal   in pls_integer
  , cstUnique  in pls_integer
  )
  is
    rec  Record_T := new_record(BRT_BEGINSST);
  begin
    write_record(rec, int2raw(cstTotal));
    write_record(rec, int2raw(cstUnique));
    put_record(stream, rec);
  end;
  
  procedure put_SSTItem (
    stream       in out nocopy stream_t
  , str          in varchar2
  , strRunArray  in StrRunArray_T default null
  )
  is
    rec  Record_T := new_record(BRT_SSTITEM);
  begin
    write_RichStr(rec, str, strRunArray);
    put_record(stream, rec);
  end;
  
  procedure put_defaultBookViews (
    stream      in out nocopy stream_t
  , firstSheet  in pls_integer
  )
  is
    rec  Record_T := new_record(BRT_BOOKVIEW);
  begin
    put_simple_record(stream, BRT_BEGINBOOKVIEWS);
    -- BrtBookView : 
    write_record(rec, '00000000');  -- xWn
    write_record(rec, '00000000');  -- yWn
    write_record(rec, '00000000');  -- dxWn
    write_record(rec, '00000000');  -- dyWn
    write_record(rec, '58020000');  -- iTabRatio : 600
    write_record(rec, int2raw(firstSheet));  -- itabFirst
    write_record(rec, int2raw(firstSheet));  -- itabCur
    write_record(rec, 
                 bitVector(
                   0  -- fHidden 
                 , 0  -- fVeryHidden
                 , 0  -- fIconic
                 , 1  -- fDspHScroll
                 , 1  -- fDspVScroll
                 , 1  -- fBotAdornment
                 , 1  -- fAFDateGroup
                 ));
    put_record(stream, rec);
    put_simple_record(stream, BRT_ENDBOOKVIEWS);
  end;
  
  procedure put_BundleSh (
    stream     in out nocopy stream_t
  , sheetId    in pls_integer
  , relId      in varchar2
  , sheetName  in varchar2
  , state      in pls_integer
  )
  is
    sheetIdx  pls_integer := sheetMap.count; -- 0-based sheet index
  begin
    sheetMap(upper(sheetName)) := sheetIdx;
    put_record(stream, make_BundleSh(sheetId, relId, sheetName, state));
  end;

  procedure put_BeginSheet (
    stream  in out nocopy stream_t
  )
  is
  begin
    put_simple_record(stream, BRT_BEGINSHEET);
    shrFmlaMap.delete;
  end;
  
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
  )
  is
  begin
    put_record(stream, make_BeginList(tableId, name, displayName, showHeader, firstRow, firstCol, lastRow, lastCol));
  end;

  procedure put_BeginListCol (
    stream     in out nocopy stream_t
  , fieldId    in pls_integer
  , fieldName  in varchar2    
  )
  is
  begin
    put_record(stream, make_BeginListCol(fieldId, fieldName));
  end;

  procedure put_BeginAFilter (
    stream    in out nocopy stream_t
  , firstRow  in pls_integer
  , firstCol  in pls_integer
  , lastRow   in pls_integer
  , lastCol   in pls_integer 
  )
  is
  begin
    put_record(stream, make_BeginAFilter(firstRow, firstCol, lastRow, lastCol));
  end;

  procedure put_ListPart (
    stream  in out nocopy stream_t
  , relId   in varchar2
  )
  is
  begin
    put_record(stream, make_ListPart(relId));
  end;
  
  procedure put_TableStyleClient (
    stream             in out nocopy stream_t
  , tableStyleName     in varchar2
  , showFirstColumn    in boolean
  , showLastColumn     in boolean
  , showRowStripes     in boolean
  , showColumnStripes  in boolean
  )
  is
  begin
    put_record( stream
              , make_TableStyleClient( tableStyleName
                                     , showFirstColumn
                                     , showLastColumn
                                     , showRowStripes
                                     , showColumnStripes ) );
  end;

  procedure put_BuiltInStyle (
    stream     in out nocopy stream_t
  , builtInId  in pls_integer
  , styleName  in varchar2
  , xfId       in pls_integer
  )
  is
  begin
    put_record(stream, make_BuiltInStyle(builtInId, styleName, xfId));
  end;
  
  procedure put_CalcProp (
    stream    in out nocopy stream_t
  , calcId    in pls_integer
  , refStyle  in pls_integer
  )
  is
  begin
    put_record(stream, make_CalcProp(calcId, refStyle));
  end;

  procedure put_WsProp (
    stream    in out nocopy stream_t
  , tabColor  in varchar2
  )
  is
  begin
    put_record(stream, make_WsProp(tabColor));
  end;

  procedure put_BeginWsView (
    stream    in out nocopy stream_t 
  , dspGrid   in boolean
  , dspRwCol  in boolean
  )
  is
  begin
    put_record(stream, make_BeginWsView(dspGrid, dspRwCol));
  end;

  procedure put_FrozenPane (
    stream   in out nocopy stream_t
  , numRows  in pls_integer
  , numCols  in pls_integer
  , topRow   in pls_integer  -- first row of the lower right pane
  , leftCol  in pls_integer  -- first column of the lower right pane
  )
  is
  begin
    put_record(stream, make_FrozenPane(numRows, numCols, topRow, leftCol));
  end;

  procedure put_NumFmt (
    stream  in out nocopy stream_t
  , id      in pls_integer
  , fmt     in varchar2
  )
  is
  begin
    put_record(stream, make_NumFmt(id, fmt));
  end;

  procedure put_Font (
    stream  in out nocopy stream_t
  , font    in ExcelTypes.CT_Font
  )
  is
  begin
    put_record(stream, make_Font(font));
  end;

  procedure put_PatternFill (
    stream       in out nocopy stream_t
  , patternFill  in ExcelTypes.CT_PatternFill
  )
  is
  begin
    put_record(stream, make_PatternFill(patternFill));
  end;

  procedure put_GradientFill (
    stream        in out nocopy stream_t
  , gradientFill  in ExcelTypes.CT_GradientFill
  )
  is
  begin
    put_record(stream, make_GradientFill(gradientFill));
  end;

  procedure put_Fill (
    stream  in out nocopy stream_t
  , fill    in ExcelTypes.CT_Fill
  )
  is
  begin
    case fill.fillType
    when ExcelTypes.FT_PATTERN then
      put_PatternFill(stream, fill.patternFill);
    when ExcelTypes.FT_GRADIENT then
      put_GradientFill(stream, fill.gradientFill);
    end case;
  end;

  procedure put_Border (
    stream  in out nocopy stream_t
  , border  in ExcelTypes.CT_Border
  )
  is
  begin
    put_record(stream, make_Border(border));
  end;

  procedure put_XF (
    stream      in out nocopy stream_t
  , xfId        in pls_integer default null
  , numFmtId    in pls_integer default 0
  , fontId      in pls_integer default 0
  , fillId      in pls_integer default 0
  , borderId    in pls_integer default 0
  , alignment   in ExcelTypes.CT_CellAlignment
  )
  is
  begin
    put_record(stream
             , make_XF(
                 xfId
               , numFmtId
               , fontId
               , fillId
               , borderId
               , alignment
               ));    
  end;

  procedure put_DXF (
    stream  in out nocopy stream_t
  , style   in ExcelTypes.CT_Style
  )
  is
  begin
    put_record(stream, make_DXF(style));    
  end;

  procedure put_ColInfo (
    stream         in out nocopy stream_t
  , colId          in pls_integer
  , colWidth       in number
  , isCustomWidth  in boolean
  , styleRef       in pls_integer
  )
  is
  begin
    put_record(stream, make_ColInfo(colId, colWidth, isCustomWidth, styleRef));
  end;

  procedure put_MergeCell (
    stream    in out nocopy stream_t
  , rwFirst   in pls_integer
  , rwLast    in pls_integer
  , colFirst  in pls_integer
  , colLast   in pls_integer
  )
  is
  begin
    put_record(stream, make_MergeCell(rwFirst, rwLast, colFirst, colLast));
  end;

  procedure put_WsFmtInfo (
    stream            in out nocopy stream_t
  , defaultRowHeight  in number
  )
  is
  begin
    put_record(stream, make_WsFmtInfo(defaultRowHeight));
  end;

  procedure put_Names (
    stream  in out nocopy stream_t
  , names   in out nocopy ExcelTypes.CT_DefinedNames
  )
  is
    externals  ExcelTypes.CT_Externals;
    recArray   recordArray_t := recordArray_t();
    idx        pls_integer;
    supLink    pls_integer;
  begin
    -- Generate BrtName records
    -- Using an iterator pattern here since a name might generate a hidden future-function name,
    -- whose corresponding record would have to be created in the same loop
    idx := names.first;
    while idx is not null loop
      recArray.extend;
      recArray(idx) := make_Name(names, idx);
      idx := names.next(idx);
    end loop;
    
    externals := ExcelFmla.getExternals;
    
    if externals.xtiArray.count != 0 then
      put_simple_record(stream, BRT_BEGINEXTERNALS);
      
      -- supporting links
      supLink := externals.supLinks.first;
      while supLink is not null loop
        put_simple_record(stream, supLink);
        supLink := externals.supLinks.next(supLink);
      end loop;

      -- Update initial sheet indices as they might have shifted due to pageable sheets expansion
      -- At this point, sheetMap should contain the correct indices
      for i in 1 .. externals.xtiArray.count loop
        if externals.xtiArray(i).firstSheet.idx != -2 then
          externals.xtiArray(i).firstSheet.idx := sheetMap(upper(externals.xtiArray(i).firstSheet.name));
          externals.xtiArray(i).lastSheet.idx := sheetMap(upper(externals.xtiArray(i).lastSheet.name));
        end if;
      end loop;
      
      -- 2.4.665 BrtExternSheet
      put_record(stream, make_ExternSheet(externals.xtiArray));
      
      put_simple_record(stream, BRT_ENDEXTERNALS);
    end if;
    
    -- Put BrtName records
    for i in 1 .. recArray.count loop
      put_record(stream, recArray(i));
    end loop;
    
  end;

  procedure put_CellImage (
    stream    in out nocopy stream_t
  , colIndex  in pls_integer
  , styleRef  in pls_integer default 0
  , vmId      in pls_integer
  )
  is
    rec  record_t;
  begin
    -- 2.4.845 BrtValueMeta
    put_simple_record(stream, BRT_VALUEMETA, int2raw(vmId));
    -- 2.4.318 BrtCellError
    rec := make_Cell(BRT_CELLERROR, colIndex, styleRef);
    write_record(rec, FT_ERR_VALUE);
    put_record(stream, rec);
  end;
  
  procedure put_CellFmla (
    stream    in out nocopy stream_t
  , colIndex  in pls_integer
  , styleRef  in pls_integer default 0
  , expr      in varchar2
  , shared    in boolean
  , si        in pls_integer default null
  , cellRef   in varchar2 default null
  , refStyle  in pls_integer default null
  )
  is
    makeShrFmla  boolean := false;
  begin
    
    if shared then
      if not shrFmlaMap.exists(si) then
        shrFmlaMap(si).ptgExp := ExcelFmla.getPtgExp(cellRef);
        makeShrFmla := true;
      end if;
      put_record(stream, make_Fmla(colIndex, styleRef, expr, si, cellRef, refStyle));
      if makeShrFmla then
        -- save a pointer to the beginning of this BrtShrFmla record so that we can set its rfx structure later
        -- when the range is known
        shrFmlaMap(si).pos := stream.sz + 1;
        put_record(stream, make_ShrFmla(expr, cellRef, refStyle));
      end if;
    else
      put_record(stream, make_Fmla(colIndex, styleRef, expr, si, cellRef, refStyle));
    end if;
  end;

  procedure put_ShrFmlaRfX (
    stream    in out nocopy stream_t
  , si        in pls_integer
  , firstRow  in pls_integer
  , firstCol  in pls_integer
  , lastRow   in pls_integer
  , lastCol   in pls_integer  
  )
  is
  begin
    seek(stream, shrFmlaMap(si).pos);
    next_record(stream);
    expect(stream, BRT_SHRFMLA);
    dbms_lob.write(stream.content, 16, stream.offset, make_RfX(firstRow, firstCol, lastRow, lastCol));
  end;

  procedure put_DVals (
    stream    in out nocopy stream_t
  , dvRules   in ExcelTypes.CT_DataValidations
  )
  is
  begin
    put_record(stream, make_BeginDVals(dvRules.count));
    for i in 1 .. dvRules.count loop
      put_record(stream, make_DVal(dvRules(i)));
    end loop;
    put_simple_record(stream, BRT_ENDDVALS);
  end;

  procedure put_CFRule (
    stream  in out nocopy stream_t
  , rule    in ExcelTypes.CT_CfRule
  , pos     in pls_integer
  )
  is
  begin
    put_record(stream, make_BeginCFRule(rule, pos)); -- BrtBeginCFRule
    
    case rule.type
    when ExcelTypes.CF_TYPE_GRADIENT then
      
      put_simple_record(stream, BRT_BEGINCOLORSCALE);
      
      -- BrtCFVO
      for i in 1 .. rule.cfvoList.count loop
        put_record(stream, make_CFVO(rule.cfvoList(i), rule.type));
      end loop;
      
      -- BrtColor
      for i in 1 .. rule.cfvoList.count loop
        put_simple_record(stream, BRT_COLOR, make_BrtColor(rule.cfvoList(i).color));
      end loop;      
      
      put_simple_record(stream, BRT_ENDCOLORSCALE);
    
    when ExcelTypes.CF_TYPE_DATABAR then
      
      put_record(stream, make_BeginDatabar(not rule.hideValue));
      
      put_record(stream, make_CFVO(rule.cfvoList(1), rule.type)); -- minimum value
      put_record(stream, make_CFVO(rule.cfvoList(2), rule.type)); -- maximum value
      put_simple_record(stream, BRT_COLOR, make_BrtColor(rule.cfvoList(3).color)); -- bar color
    
      put_simple_record(stream, BRT_ENDDATABAR);
    
    when ExcelTypes.CF_TYPE_MULTISTATE then
      
      put_record(stream, make_BeginIconSet(rule.iconSet, rule.hideValue, rule.reverse));
      
      -- BrtCFVO
      -- first CFVO must be a dummy one, which appears to default to this : 
      put_record(stream, make_CFVO(ExcelTypes.makeCfvo(ExcelTypes.CFVO_PERCENT, 0, true), rule.type));
      for i in 1 .. rule.cfvoList.count loop
        put_record(stream, make_CFVO(rule.cfvoList(i), rule.type));
      end loop;      
      
      put_simple_record(stream, BRT_ENDICONSET);
    
    else
      null;
    end case;
    
    put_simple_record(stream, BRT_ENDCFRULE); -- BrtEndCFRule
  end;

  procedure put_CondFmts (
    stream    in out nocopy stream_t
  , cfRules   in ExcelTypes.CT_CfRules
  )
  is
  begin
    for i in 1 .. cfRules.count loop
      put_record(stream, make_BeginCondFormat(cfRules(i))); -- BrtBeginConditionalFormatting
      put_CFRule(stream, cfRules(i), i);
      put_simple_record(stream, BRT_ENDCONDFORMAT); -- BrtEndConditionalFormatting
    end loop;
  end;

  procedure put_Metadata (
    stream      in out nocopy stream_t
  , imageCount  in pls_integer
  )
  is
    metadataType  varchar2(256) := 'XLRICHVALUE';
  begin
    put_simple_record(stream, BRT_BEGINMETADATA);
    
    -- 2.4.75 BrtBeginEsmdtinfo
    put_simple_record(stream, BRT_BEGINESMDTINFO, int2raw(1)); -- cMdtinfo = 1
    put_record(stream, make_MdtInfo(metadataType));
    put_simple_record(stream, BRT_ENDESMDTINFO);
    
    -- 2.4.73 BrtBeginEsfmd
    put_record(stream, make_BeginEsfmd(imageCount, metadataType));
    for i in 1 .. imageCount loop
      put_simple_record(stream, BRT_BEGINFMD);
      put_simple_record(stream, BRT_FRTBEGIN, '01000200'); -- productVersion
      
      -- /!\ 2.4.194 BrtBeginRichValueBlock: the specs say it contains FRTHeader + irv data
      -- but actually Excel puts it in the BrtEndRichValueBlock record
      put_simple_record(stream, BRT_BEGINRICHVALUEBLOCK);
      put_simple_record(stream, BRT_ENDRICHVALUEBLOCK, utl_raw.concat('00000000', int2raw(i-1))); -- FRTHeader + irv
      
      put_simple_record(stream, BRT_FRTEND);
      put_simple_record(stream, BRT_ENDFMD);
    end loop;
    put_simple_record(stream, BRT_ENDESFMD);
    
    -- 2.4.74 BrtBeginEsmdb
    put_simple_record(stream, BRT_BEGINESMDB, utl_raw.concat(int2raw(imageCount), '00000000')); -- cMdb + fCellMeta
    for i in 1 .. imageCount loop
      -- 2.4.703 BrtMdb
      put_simple_record(stream
                      , BRT_MDB
                      , utl_raw.concat(
                          int2raw(1)    -- cMdir
                        , int2raw(1)    -- Mdir.iMdt
                        , int2raw(i-1)  -- Mdir.mdd
                        ));
    end loop;
    put_simple_record(stream, BRT_ENDESMDB);
    
    put_simple_record(stream, BRT_ENDMETADATA);
  end;

  procedure put_Drawing (
    stream  in out nocopy stream_t
  , rId     in varchar2
  )
  is
    rec record_t := new_record(BRT_DRAWING);
  begin
    write_XLWideString(rec, rId);
    put_record(stream, rec);
  end;
  
  -- 2.4.307 BrtBkHim
  procedure put_BkHim (
    stream  in out nocopy stream_t
  , rId     in varchar2
  )
  is
    rec record_t := new_record(BRT_BKHIM);
  begin
    write_XLWideString(rec, rId);
    put_record(stream, rec);
  end;

  -- convert a 0-based column number to base26 string
  function base26encode (colNum in pls_integer) 
  return varchar2
  result_cache
  is
    output  varchar2(3);
    num     pls_integer := colNum + 1;
  begin
    if colNum is not null then
      while num != 0 loop
        output := chr(65 + mod(num-1,26)) || output;
        num := trunc((num-1)/26);
      end loop;
    end if;
    return output;
  end;

  function base26decode (colRef in varchar2)
  return pls_integer
  result_cache
  is
  begin
    return ascii(substr(colRef,-1,1))-65 
         + nvl((ascii(substr(colRef,-2,1))-64)*26, 0)
         + nvl((ascii(substr(colRef,-3,1))-64)*676, 0);
  end;

  function parseColumnList (
    cols  in varchar2
  , sep   in varchar2 default ','
  )
  return ColumnMap_T
  is
    colMap  Columnmap_t;
    i       pls_integer;
    token   varchar2(3);
    p1      pls_integer := 1;
    p2      pls_integer;  
  begin
    if cols is not null then
      loop
        p2 := instr(cols, sep, p1);
        if p2 = 0 then
          token := substr(cols, p1);
        else
          token := substr(cols, p1, p2-p1);    
          p1 := p2 + 1;
        end if;
        i := base26decode(token);
        colMap(i) := token;
        exit when p2 = 0;
      end loop;
    end if;
    return colMap; 
  end;

  procedure next_sheet (
    ctx  in out nocopy Context_T
  )
  is
    has_next  boolean := (ctx.curr_sheet < ctx.sheetList.count);
  begin
    if ctx.curr_sheet > 0 then
      close_stream(ctx.stream);
    end if;
    while has_next loop
      ctx.curr_sheet := ctx.curr_sheet + 1;
      debug('Switching to sheet '||ctx.curr_sheet);
      ctx.stream := open_stream(ctx.sheetList(ctx.curr_sheet));
      exit;
    end loop;
    if not has_next then
      debug('End of sheet list');
      ctx.done := true;
    end if;
  end;

  function read_Bool (stream in out nocopy Stream_T)
  return String_T
  is
    bBool  raw(1);
    str    String_T;
  begin
    bBool := read_bytes(stream, 1);
    str.strValue := case bBool 
                      when BOOL_TRUE then 'TRUE' 
                      when BOOL_FALSE then 'FALSE' 
                    end; 
    return str;
  end;

  function read_Err (stream in out nocopy Stream_T)
  return String_T
  is
    fErr  raw(1);
    str   String_T;
  begin
    fErr := read_bytes(stream, 1);
    str.strValue := 
      case fErr
            when FT_ERR_NULL then '#NULL!'
            when FT_ERR_DIV_ZERO then '#DIV/0!'
            when FT_ERR_VALUE then '#VALUE!'
            when FT_ERR_REF then '#REF!'
            when FT_ERR_NAME then '#NAME?'
            when FT_ERR_NUM then '#NUM!'
            when FT_ERR_NA then '#N/A'
            when FT_ERR_GETDATA then '#GETTING_DATA'
          end;     
     return str;
  end;
  
  function read_RK (stream in out nocopy Stream_T)
  return number
  is
    rk  RK_T;
    nm  number;
  begin
    rk.RkNumber := read_bytes(stream, 4);
    rk.fX100 := is_bit_set(utl_raw.substr(rk.RkNumber,1,1), 0);
    rk.fInt := is_bit_set(utl_raw.substr(rk.RkNumber,1,1), 1);
    
    rk.RkNumber := utl_raw.bit_and(rk.RkNumber, 'FCFFFFFF');
    if rk.fInt then 
      -- convert to int and shift right 2
      nm := to_number(utl_raw.cast_to_binary_integer(rk.RkNumber, utl_raw.little_endian)/4);
    else
      -- pad LSBs with 0 and convert to double
      nm := to_number(utl_raw.cast_to_binary_double(utl_raw.concat('00000000',rk.RkNumber),utl_raw.little_endian));
    end if;
    if rk.fX100 then
      nm := nm/100;
    end if;
    
    return nm;
  end;
  
  function read_Number (stream in out nocopy Stream_T)
  return number
  is
    Xnum  raw(8);
  begin
    Xnum := read_bytes(stream, 8);
    return to_number(utl_raw.cast_to_binary_double(Xnum, utl_raw.little_endian));
  end;

  function read_SSTItem (
    stream  in out nocopy Stream_T
  , sst     in SST_T 
  )
  return String_T
  is
    idx  pls_integer := read_int32(stream) + 1;
  begin
    return sst.strings(idx);
  end;

  procedure read_XLString (
    stream  in out nocopy Stream_T
  , str     in out nocopy XLRichString_T
  , sttype  in pls_integer default ST_SIMPLE
  )
  is
    raw1    raw(1);
    buf     raw(32764);
    cbuf    varchar2(32764);
    rem     pls_integer;
    amt     pls_integer;
    csz     pls_integer := 2;
    csname  varchar2(30) := 'AL16UTF16LE';
  begin
    
    if stType = ST_RICHSTR then
      raw1 := read_bytes(stream, 1);
      str.fRichStr := is_bit_set(raw1, 0);
      str.fExtStr := is_bit_set(raw1, 1);
    end if;
    
    str.cch := read_int32(stream);
    
    if str.cch != -1 then
    
      rem := str.cch; -- characters left to read;
      
      while rem != 0 loop
        
        amt := least(8191, rem) * csz; -- byte amount to read      
        buf := read_bytes(stream, amt);
        cbuf := utl_i18n.raw_to_char(buf, csname);
        
        if str.content.is_lob then
        
          str.content.lobValue := str.content.lobValue || cbuf;
        
        else
          
          str.byte_len := str.byte_len + lengthb(cbuf);
          if str.byte_len > MAX_STRING_SIZE then
            -- switch to lob storage
            dbms_lob.createtemporary(str.content.lobValue, true);
            if str.content.strValue is not null then
              dbms_lob.writeappend(str.content.lobValue, length(str.content.strValue), str.content.strValue);
            end if;
            dbms_lob.writeappend(str.content.lobValue, length(cbuf), cbuf);
            str.content.is_lob := true;
            str.content.strValue := null;
          else       
            str.content.strValue := str.content.strValue || cbuf;        
          end if;  
        
        end if;
        
        rem := rem - amt/csz;
        
      end loop;
    
    end if;
        
  end;
  
  function read_XLString (
    stream  in out nocopy Stream_T
  , stType  in pls_integer default 0
  )
  return String_T
  is
    xlstr  XLRichString_T;
  begin
    read_XLString(stream, xlstr, stType);
    return xlstr.content;
  end;
  
  function read_SheetInfo (
    stream  in out nocopy Stream_T
  )
  return BrtBundleSh_T
  is
    sh  BrtBundleSh_T;
  begin
    sh.hsState := read_bytes(stream, 4);
    sh.iTabID := read_bytes(stream, 4);
    sh.strRelID := read_XLString(stream).strValue;
    debug(sh.strRelID);
    sh.strName := read_XLString(stream).strValue;
    debug(sh.strName);
    return sh;
  end;
  
  procedure read_all (file in blob) is
    stream  Stream_T;
    --raw1    raw(1);
  begin
  
    loadRecordTypeLabels;
  
    stream := open_stream(file);

    while stream.offset < stream.sz loop
      next_record(stream);
      
      dbms_output.put_line(
        utl_lms.format_message(
          '%s %s %s' || case when stream.rsize != 0 then ' start=0x%s size=%d' end
        , lpad(stream.rt,4)
        , rpad('['||rawtohex(make_rt(stream.rt))||']',6)
        , rpad(case when recordTypeLabelMap.exists(stream.rt) then recordTypeLabelMap(stream.rt) else '<Unknown>' end, 34)
        , to_char(stream.rstart - 1, 'fm0XXXXXXXXX')
        , stream.rsize
        )
      );
      
      /*
      if stream.rt = BRT_XF then
        dbms_output.put_line('ixfeParent='||read_bytes(stream,2));
        dbms_output.put_line('iFmt='||raw2int(read_bytes(stream,2)));
        dbms_output.put_line('iFont='||raw2int(read_bytes(stream,2)));
        dbms_output.put_line('iFill='||raw2int(read_bytes(stream,2)));
        dbms_output.put_line('ixBorder='||raw2int(read_bytes(stream,2)));
        dbms_output.put_line('trot='||raw2int(read_bytes(stream,1)));
        dbms_output.put_line('indent='||raw2int(read_bytes(stream,1)));
        raw1 := read_bytes(stream,1);
        dbms_output.put_line('alc='||raw2int(utl_raw.bit_and(raw1,'07')));
        dbms_output.put_line('alcv='||raw2int(utl_raw.bit_and(raw1,'38'))/8);
        dbms_output.put_line('fWrap='||raw2int(utl_raw.bit_and(raw1,'40'))/64);
        dbms_output.put_line('fJustLast='||raw2int(utl_raw.bit_and(raw1,'80'))/128);
        raw1 := read_bytes(stream,1);
        dbms_output.put_line('fShrinkToFit='||raw2int(utl_raw.bit_and(raw1,'01')));
        dbms_output.put_line('fMergeCell='||raw2int(utl_raw.bit_and(raw1,'02'))/2);
        dbms_output.put_line('iReadingOrder='||raw2int(utl_raw.bit_and(raw1,'0C'))/4);
        dbms_output.put_line('fLocked='||raw2int(utl_raw.bit_and(raw1,'10'))/16);
        dbms_output.put_line('fHidden='||raw2int(utl_raw.bit_and(raw1,'20'))/32);
        dbms_output.put_line('fSxButton='||raw2int(utl_raw.bit_and(raw1,'40'))/64);
        dbms_output.put_line('f123Prefix='||raw2int(utl_raw.bit_and(raw1,'80'))/128);
        raw1 := read_bytes(stream,1);
        dbms_output.put_line('xfGrbitAtr.0='||raw2int(utl_raw.bit_and(raw1,'01')));
        dbms_output.put_line('xfGrbitAtr.1='||raw2int(utl_raw.bit_and(raw1,'02'))/2);
        dbms_output.put_line('xfGrbitAtr.2='||raw2int(utl_raw.bit_and(raw1,'04'))/4);
        dbms_output.put_line('xfGrbitAtr.3='||raw2int(utl_raw.bit_and(raw1,'08'))/8);
        dbms_output.put_line('xfGrbitAtr.4='||raw2int(utl_raw.bit_and(raw1,'10'))/16);
        dbms_output.put_line('xfGrbitAtr.5='||raw2int(utl_raw.bit_and(raw1,'20'))/32);
      end if;
      */
        
    end loop;
    close_stream(stream);
  end;

  function get_sheetEntries (
    p_workbook  in blob
  )
  return SheetEntries_T
  is
    i             pls_integer := 0;
    stream        Stream_T;
    sh            BrtBundleSh_T;
    sheetEntries  SheetEntries_T := SheetEntries_T();
  begin

    stream := open_stream(p_workbook);

    next_record(stream);
    -- read records until BrtEndBundleShs is found
    while stream.rt != BRT_ENDBUNDLESHS loop    
      if stream.rt = BRT_BUNDLESH then
        i := i + 1;
        sheetEntries.extend;
        sh := read_SheetInfo(stream);
        sheetEntries(i).name := sh.strName;
        sheetEntries(i).relId := sh.strRelID;
      end if;   
      next_record(stream);
    end loop;
    
    close_stream(stream);
    
    return sheetEntries;
    
  end;

  function read_CellBlock (
    ctx    in out nocopy Context_T
  , nrows  in pls_integer
  )
  return ExcelTableCellList
  is
    rcnt       pls_integer := 1;
    rw         pls_integer := ctx.curr_rw;
    col        pls_integer;
    num        number;
    str        String_T;
    emptyRow   boolean := true;
    xfId       pls_integer;
    dateXfMap  dateXfMap_t := ctx.dtStyles;
    
    cells  ExcelTableCellList := ExcelTableCellList();
    
    procedure read_col is
    begin
      col := read_int32(ctx.stream);
      xfId := utl_raw.cast_to_binary_integer(
                utl_raw.bit_and(read_bytes(ctx.stream, 4), 'FFFFFF00') -- iStyleRef (24 bits)
              , utl_raw.little_endian
              );
      --skip(ctx.stream, 4); -- iStyleRef, fPhShow, reserved
    end;
    
    function convertNumData return anydata is
    begin
      if dateXfMap.exists(xfId) then
        if dateXfMap(xfId).isTimestamp then
          return anydata.ConvertTimestamp(ExcelTypes.fromOADate(num,3));
        else
          return anydata.ConvertDate(ExcelTypes.fromOADate(num));
        end if;
      else 
        return anydata.ConvertNumber(num);
      end if;      
    end;  
    
    function get_comment (cell_ref in varchar2) return varchar2 is
    begin
      if ctx.comments.exists(ctx.curr_sheet) and ctx.comments(ctx.curr_sheet).exists(cell_ref) then
        return ctx.comments(ctx.curr_sheet)(cell_ref);
      else
        return null;
      end if;
    end;
    
    procedure add_cell(val in anydata) is
    begin
      if ctx.rng.colMap.count = 0 or ctx.rng.colMap.exists(col) then
        cells.extend;
        cells(cells.last) := new ExcelTableCell(rw + 1, base26encode(col), null, val, ctx.curr_sheet, get_comment(base26encode(col) || to_char(rw + 1)));
        emptyRow := false;
      end if;
    end;
    
  begin
    /*
    if rw is not null then
      rcnt := 1;
    end if;
    */
    next_record(ctx.stream);
    
    loop
           
      if ctx.stream.rt = BRT_ENDSHEETDATA then
        next_sheet(ctx);
        
      elsif ctx.stream.rt = BRT_ROWHDR then
        rw := read_int32(ctx.stream);
        debug('Row '||rw);
        
        if rw > ctx.rng.lastRow then
          debug('End of range');
          next_sheet(ctx);
        
        elsif rw >= ctx.rng.firstRow then
          
          -- if previous row was not empty
          if not emptyRow then
            rcnt := rcnt + 1;
          end if;
          --rcnt := rcnt + 1;
          
          if rcnt > nrows then
            ctx.curr_rw := rw;
            exit;
          end if;
          
          emptyRow := true;
          
        end if;
        
      elsif rw >= ctx.rng.firstRow then
        
        case ctx.stream.rt
        when BRT_CELLBLANK then
        
          read_col;
        
        when BRT_CELLRK    then
          
          read_col;
          num := read_RK(ctx.stream);          
          add_cell(convertNumData);
        
        when BRT_CELLERROR then
          
          read_col;
          str := read_Err(ctx.stream);
          add_cell(anydata.ConvertVarchar2(str.strValue));
        
        when BRT_CELLBOOL  then
          
          read_col;
          str := read_Bool(ctx.stream);
          add_cell(anydata.ConvertVarchar2(str.strValue));
        
        when BRT_CELLREAL  then
          
          read_col;
          num := read_Number(ctx.stream);
          add_cell(convertNumData);
          
        when BRT_CELLST    then
          
          read_col;
          str := read_XLString(ctx.stream);
          if str.is_lob then
            add_cell(anydata.ConvertClob(str.lobValue));
          else
            add_cell(anydata.ConvertVarchar2(str.strValue));
          end if;
        
        when BRT_CELLISST  then
          
          read_col;
          str := read_SSTItem(ctx.stream, ctx.sst);
          if str.is_lob then
            add_cell(anydata.ConvertClob(str.lobValue));
          else
            add_cell(anydata.ConvertVarchar2(str.strValue));
          end if;          
          
        when BRT_FMLASTRING  then
          
          read_col;
          str := read_XLString(ctx.stream);
          if str.is_lob then
            add_cell(anydata.ConvertClob(str.lobValue));
          else
            add_cell(anydata.ConvertVarchar2(str.strValue));
          end if;
        
        when BRT_FMLANUM  then
        
          read_col;
          num := read_Number(ctx.stream);
          add_cell(convertNumData);
        
        when BRT_FMLABOOL  then
        
          read_col;
          str := read_Bool(ctx.stream);
          add_cell(anydata.ConvertVarchar2(str.strValue));
        
        when BRT_FMLAERROR  then
        
          read_col;
          str := read_Err(ctx.stream);
          add_cell(anydata.ConvertVarchar2(str.strValue));
          
        else
          null;
        end case;
      
      end if;
      
      exit when ctx.done;
      
      next_record(ctx.stream);
    
    end loop;
    
    debug('cells.count='||cells.count);
    
    return cells;
    
  end;

  procedure read_SST (
    sst_part  in blob
  , sst       in out nocopy SST_T
  )
  is
    stream  Stream_T;
  begin
    stream := open_stream(sst_part);
    
    next_record(stream);
    expect(stream, BRT_BEGINSST);
    sst.cstTotal := read_int32(stream);
    sst.cstUnique := read_int32(stream);
    
    debug('sst.cstTotal = '||sst.cstTotal);
    debug('sst.cstUnique = '||sst.cstUnique);
    
    sst.strings := String_Array_T();
    sst.strings.extend(sst.cstUnique);
    
    for i in 1 .. sst.cstUnique loop
      next_record(stream);
      sst.strings(i) := read_XLString(stream, ST_RICHSTR);
      --debug(sst.strings(i).strValue);
    end loop;
    
    close_stream(stream);
  
  end;

  function read_Comment (
    stream  in out nocopy Stream_T 
  )
  return Comment_T
  is
    cmt  Comment_T;
  begin
    skip(stream, 4); -- iauthor
    cmt.rw := read_int32(stream); -- rwFirst
    skip(stream, 4); -- rwLast
    cmt.col := read_int32(stream); -- colFirst
    skip(stream, 4); -- colLast
    next_record(stream);
    expect(stream, BRT_COMMENTTEXT);
    cmt.text := read_XLString(stream, ST_RICHSTR).strValue;
    --debug(cmt.text);
    return cmt;
  end;

  function read_Comments (
    content  in blob
  )
  return CommentMap_T
  is
    stream      Stream_T;
    commentMap  CommentMap_T;
    cmt         Comment_T;
  begin
    stream := open_stream(content);
    next_record(stream);
    expect(stream, BRT_BEGINCOMMENTS);
    
    -- read records until BrtEndCommentList is found
    while stream.rt != BRT_ENDCOMMENTLIST loop    
      if stream.rt = BRT_BEGINCOMMENT then
        cmt := read_Comment(stream);
        commentMap(base26encode(cmt.col) || to_char(cmt.rw + 1)) := cmt.text;
      end if;
      next_record(stream);
    end loop;
    
    close_stream(stream); 
    return commentMap;

  end;
  
  function read_Styles (
    content  in blob
  )
  return dateXfMap_t
  is
    stream      Stream_T;
    cnt         pls_integer;
    numFmtId    pls_integer;
    numFmt      ExcelTypes.CT_NumFmt;
    numFmtMap   ExcelTypes.CT_NumFmtMap := ExcelTypes.getBuiltInDateFmts();
    dateXfMap   dateXfMap_t;
    raw1        raw(1);
    xfId        pls_integer := 0;
  begin
    stream := open_stream(content);
    next_record(stream);
    expect(stream, BRT_BEGINSTYLESHEET);
    
    -- read BrtFmt records
    seek_first(stream, BRT_BEGINFMTS);
    if stream.rt = BRT_BEGINFMTS then
      cnt := read_int32(stream);
      if cnt != 0 then

        while stream.rt != BRT_ENDFMTS loop    
          if stream.rt = BRT_FMT then
            numFmt := ExcelTypes.makeNumFmt(
                        numFmtId   => read_int16(stream)
                      , formatCode => read_XLString(stream).strValue
                      );
            if numFmt.isDate then
              numFmtMap(numFmt.numFmtId) := numFmt;
            end if;
          end if;
          next_record(stream);
        end loop;
        
        if numFmtMap.count != 0 then
          -- read BrtXf records
          seek_first(stream, BRT_BEGINCELLXFS);
          while stream.rt != BRT_ENDCELLXFS loop    
            if stream.rt = BRT_XF then
              skip(stream, 2); -- ixfeParent
              numFmtId := read_int16(stream); -- iFmt
              skip(stream, 4); -- iFont, iFill
              skip(stream, 4); -- ixBorder, trot, indent
              skip(stream, 2); -- alc, alcv, A-I
              raw1 := read_bytes(stream, 1); -- xfGrbitAtr
              if is_bit_set(raw1, 0) and numFmtMap.exists(numFmtId) then
                dateXfMap(xfId) := numFmtMap(numFmtId);
              end if;
              xfId := xfId + 1;
            end if;
            next_record(stream);
          end loop;          
        
        end if;
                
      end if;
    end if;
    close_stream(stream);
    return dateXfMap;
  end;
  
  procedure resetSheetCache is
  begin
    sheetMap.delete;
  end;
  
  function new_context (
    p_sst_part  in blob
  , p_cols      in varchar2 default null
  , p_firstRow  in pls_integer default null
  , p_lastRow   in pls_integer default null
  , p_styles    in blob default null
  )
  return pls_integer
  is
    ctx     Context_T;
    ctx_id  pls_integer; 
  begin
      
    if p_sst_part is not null then
      read_sst(p_sst_part, ctx.sst);
    end if;
    
    if p_styles is not null then
      ctx.dtStyles := read_Styles(p_styles);
    end if;
  
    ctx.rng.firstRow := nvl(p_firstRow, 1) - 1;
    ctx.rng.lastRow := nvl(p_lastRow, 1048576) - 1;
    ctx.rng.colMap := parseColumnList(p_cols);
    ctx.done := false;
    ctx.sheetList := SheetList_T();
    ctx.curr_sheet := 0;
    
    ctx_id := nvl(ctx_cache.last, 0) + 1;
    ctx_cache(ctx_id) := ctx;
    
    return ctx_id;
    
  end;

  procedure add_sheet (
    p_ctx_id    in pls_integer
  , p_content   in blob
  , p_comments  in blob
  )
  is
    i  pls_integer;
  begin
    ctx_cache(p_ctx_id).sheetList.extend;
    i := ctx_cache(p_ctx_id).sheetList.last;
    ctx_cache(p_ctx_id).sheetList(i) := p_content;
    -- comments
    if p_comments is not null then
      ctx_cache(p_ctx_id).comments(i) := read_Comments(p_comments);
    end if;
  end;

  function iterate_context (
    p_ctx_id  in pls_integer
  , p_nrows   in pls_integer
  )
  return ExcelTableCellList
  is
    cells  ExcelTableCellList;
  begin
    if ctx_cache(p_ctx_id).curr_sheet = 0 then
      next_sheet(ctx_cache(p_ctx_id));
    end if;
    if not ctx_cache(p_ctx_id).done then
      cells := read_CellBlock(ctx_cache(p_ctx_id), p_nrows);
    end if;
    return cells;
  end;

  procedure free_context (
    p_ctx_id  in pls_integer 
  )
  is
  begin
    --close_stream(ctx_cache(p_ctx_id).stream);
    ctx_cache(p_ctx_id).sst.strings := String_Array_T();
    ctx_cache.delete(p_ctx_id);
  end;

  procedure read_formulas (file in blob) is
    stream      Stream_T;
    recordName  varchar2(128);
    intValue    pls_integer;
    
    byte1       raw(1);
    byte2       raw(1);
    raw1        raw(1);
    readCount   pls_integer;
    cce         pls_integer;
    cb          pls_integer;
    int1        pls_integer;
    int2        pls_integer;
    
    rgceOffset  pls_integer;
    
    rc          sys_refcursor;
    type ptg_t is record (id pls_integer, name varchar2(128), sz pls_integer, hasType pls_integer);
    type ptgList_t is table of ptg_t index by pls_integer;
    
    type ptgExtraList_t is table of varchar2(128);
    extras  ptgExtraList_t := ptgExtraList_t();
    
    ptgList  ptgList_t;
    ptg      ptg_t;
    ptgType  pls_integer;
    
    cnt  pls_integer;
    
    type uncheckedRfX_t is record (rwFirst pls_integer, rwLast pls_integer, colFirst pls_integer, colLast pls_integer);
    rfx  uncheckedRfX_t;
    
    dvFmla   boolean;
    
  begin
    
    open rc for 'select utl_raw.cast_to_binary_integer(utl_raw.concat(byte1,byte2)), ptg_name, ptg_size, decode(has_type,''Y'',1,0) from xlsb_ptg';
    loop
      fetch rc into ptg.id, ptg.name, ptg.sz, ptg.hasType;
      exit when rc%notfound;
      ptgList(ptg.id) := ptg; 
    end loop;
    close rc;
    
    loadRecordTypeLabels;
    stream := open_stream(file);

    while stream.offset < stream.sz loop
      next_record(stream);   
      continue when stream.rt not in (BRT_FMLASTRING,BRT_FMLANUM,BRT_FMLABOOL,BRT_FMLAERROR,BRT_ARRFMLA,BRT_NAME,BRT_SHRFMLA,BRT_DVAL);
      
      debug('---------------------------');
      debug(recordTypeLabelMap(stream.rt));
      debug('---------------------------');
      
      dvFmla := false;
      
      if stream.rt = BRT_ARRFMLA then

        skip(stream, 16); -- rfx
        skip(stream, 1); -- A + unused
        
      elsif stream.rt = BRT_SHRFMLA then
        
        skip(stream, 16); -- rfx
        
      elsif stream.rt = BRT_NAME then
        
        raw1 := read_bytes(stream, 1);
        debug('fHidden='||case when is_bit_set(raw1,0) then 1 else 0 end);
        debug('fFunc='||case when is_bit_set(raw1,1) then 1 else 0 end);
        debug('fOB='||case when is_bit_set(raw1,2) then 1 else 0 end);
        debug('fProc='||case when is_bit_set(raw1,3) then 1 else 0 end);
        debug('fCalcExp='||case when is_bit_set(raw1,4) then 1 else 0 end);
        debug('fBuiltin='||case when is_bit_set(raw1,5) then 1 else 0 end);
        raw1 := read_bytes(stream, 1);
        debug('fPublished='||case when is_bit_set(raw1,7) then 1 else 0 end);
        raw1 := read_bytes(stream, 1);
        debug('fWorkbookParam='||case when is_bit_set(raw1,0) then 1 else 0 end);
        debug('fFutureFunction='||case when is_bit_set(raw1,1) then 1 else 0 end);
        skip(stream, 1); -- ...reserved
        skip(stream, 1); -- chKey
        debug('itab='||read_int32(stream)); -- itab
        debug('name='||read_XLString(stream).strValue); -- name
        
      elsif stream.rt = BRT_DVAL then
        
        dvFmla := true;
        
        skip(stream, 4); -- valType .. reserved
        cnt := read_int32(stream); -- sqrfx.crfx
        skip(stream, cnt * 16); -- rgrfx
        cnt := read_int32(stream); -- DValStrings.strErrorTitle.cchCharacters
        if cnt != -1 then
           skip(stream, cnt * 2);     -- DValStrings.strErrorTitle.rgchData
        end if;
        cnt := read_int32(stream); -- DValStrings.strError.cchCharacters
        if cnt != -1 then
          skip(stream, cnt * 2);     -- DValStrings.strError.rgchData
        end if;
        cnt := read_int32(stream); -- DValStrings.strPromptTitle.cchCharacters
        if cnt != -1 then
          skip(stream, cnt * 2);     -- DValStrings.strPromptTitle.rgchData
        end if;
        cnt := read_int32(stream); -- DValStrings.strPrompt.cchCharacters
        if cnt != -1 then
          skip(stream, cnt * 2);     -- DValStrings.strPrompt.rgchData
        end if;
        
      else
      
        skip(stream, 8); -- Cell
                
        case stream.rt
        when BRT_FMLASTRING then
          intValue := read_int32(stream); -- cchCharacters
          skip(stream, intValue * 2); -- rgchData
        when BRT_FMLANUM then
          skip(stream, 8); -- xnum
        when BRT_FMLABOOL then
          skip(stream, 1); -- bBool
        when BRT_FMLAERROR then
          skip(stream, 1); -- fErr
        end case;
          
        skip(stream, 2); -- grbitFlags
      
      end if;
      
      <<READ_FORMULA>>
      debug('BEGIN ParsedFormula');
      
      extras.delete;
      cce := read_int32(stream); -- cce
      debug('cce='||cce);
      
      debug('BEGIN Rgce ['||to_char(stream.offset-1,'FM0XXXXXXX')||']');
      
      rgceOffset := 0;
      
      while rgceOffset < cce loop
        
        byte1 := read_bytes(stream, 1); -- ptg byte1
        readCount := 1;
        if byte1 in ('18','19') then
          byte2 := read_bytes(stream, 1); -- read byte2
          readCount := readCount + 1;
        else
          byte2 := null;
        end if;
        ptg := ptgList(utl_raw.cast_to_binary_integer(utl_raw.concat(byte1,byte2)));
        
        -- if that ptg requires an extra
        if ptg.name in ('PtgArray','PtgMemArea','PtgExp') then
          extras.extend;
          extras(extras.last) := ptg.name;
        end if;
        
        if ptg.hasType = 1 then
          -- select bits 5 and 6, and shift right 5
          ptgType := raw2int(utl_raw.bit_and(byte1, '60'))/32;
        else
          ptgType := null;
        end if;
        
        debug(utl_lms.format_message('[%s]%s', ptg.name
                                               , case ptgType 
                                                 when 1 then '[REFERENCE]'
                                                 when 2 then '[VALUE]'
                                                 when 3 then '[ARRAY]'
                                                 end
                                                   ));
        
        if ptg.sz is not null then
          skip(stream, ptg.sz - readCount);
        else
          
          if ptg.name = 'PtgStr' then
            intValue := read_int16(stream);  -- cch
            skip(stream, intValue*2); -- rgch
            ptg.sz := 3 + intValue*2;
          elsif ptg.name = 'PtgAttrChoose' then
            intValue := read_int16(stream); -- cOffset
            skip(stream, (intValue+1)*2); -- rgOffset
            ptg.sz := 4 + (intValue+1)*2;
          end if;
        
        end if;
        
        rgceOffset := rgceOffset + ptg.sz;
      
      end loop;
      
      debug('END Rgce');
      
      debug('BEGIN RgbExtra ['||to_char(stream.offset-1,'FM0XXXXXXX')||']');
      cb := read_int32(stream); -- cb
      debug('cb='||cb);
      
      if cb != 0 then
        
        for i in 1 .. extras.count loop
           
          case extras(i)
          when 'PtgMemArea' then
            -- read PtgExtraMem
            debug('[PtgExtraMem]');
            cnt := read_int32(stream); -- count
            for j in 1 .. cnt loop
              rfx.rwFirst := read_int32(stream);
              rfx.rwLast := read_int32(stream);
              rfx.colFirst := read_int32(stream);
              rfx.colLast := read_int32(stream);
              debug(
                '  '||
                base26encode(rfx.colFirst) ||
                to_char(rfx.rwFirst+1) || ':' ||
                base26encode(rfx.colLast) ||
                to_char(rfx.rwLast+1)
              );
            end loop;
            
          when 'PtgExp' then
            --read PtgExtraCol
            debug('[PtgExtraCol]');
            debug('  '||base26encode(read_int32(stream)));
            
          when 'PtgArray' then
            debug('[PtgExtraArray]');
            int1 := read_int32(stream);
            int2 := read_int32(stream);
            debug('  rows='||int1);
            debug('  cols='||int2);
            for j in 1 .. int1 * int2 loop
              raw1 := read_bytes(stream, 1);
              case raw1
              when '00' then
                debug('  SerNum');
                skip(stream, 8);
              when '01' then
                debug('  SerStr');
                int1 := read_int16(stream);
                skip(stream, int1 * 2);
              when '02' then
                debug('  SerBool');
                skip(stream, 1);
              when '04' then
                debug('  SerErr');
                skip(stream, 4);
              end case;
            end loop;
          
          end case;
        
        end loop;
        --skip(stream, cb); -- rgcb
        
      end if;
      
      debug('END RgbExtra');
      
      debug('END ParsedFormula');
      
      -- if this is a DVParsedFormula, iterate to read formula2
      if dvFmla then
        dvFmla := false;
        goto READ_FORMULA;
      end if;
        
    end loop;
    close_stream(stream);
  end;

begin
  
  init();

end xutl_xlsb;
/
