create or replace package body xutl_xlsb is
 
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
  BRT_NAME              constant pls_integer := 39;
  BRT_FONT              constant pls_integer := 43;
  BRT_FMT               constant pls_integer := 44;
  BRT_FILL              constant pls_integer := 45;
  BRT_BORDER            constant pls_integer := 46;
  BRT_XF                constant pls_integer := 47;
  BRT_STYLE             constant pls_integer := 48;
  BRT_COLINFO           constant pls_integer := 60;
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
  BRT_BEGINLIST         constant pls_integer := 343;
  BRT_BEGINLISTCOL      constant pls_integer := 347;
  BRT_BEGINCOLINFOS     constant pls_integer := 390;
  BRT_ENDCOLINFOS       constant pls_integer := 391;
  BRT_EXTERNSHEET       constant pls_integer := 362;
  BRT_WSFMTINFO         constant pls_integer := 485;
  BRT_TABLESTYLECLIENT  constant pls_integer := 513;
  BRT_BEGINCOMMENTS     constant pls_integer := 628;
  BRT_ENDCOMMENTLIST    constant pls_integer := 634;
  BRT_BEGINCOMMENT      constant pls_integer := 635;
  BRT_COMMENTTEXT       constant pls_integer := 637;
  BRT_LISTPART          constant pls_integer := 661;
  
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
  
  type Context_T is record (
    stream      Stream_T
  , sst         SST_T
  , rng         Range_T
  , curr_rw     pls_integer
  , done        boolean
  , sheetList   SheetList_T
  , curr_sheet  pls_integer
  , comments    Comments_T
  );
  
  type Context_cache_T is table of Context_T index by pls_integer;
  ctx_cache  Context_cache_T;
  
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
    rColorCode  raw(4) := hextoraw(colorCode);
  begin
    if colorCode is not null then
      return utl_raw.concat( '05'    -- fValidRGB = 1, xColorType = 2
                           , '00'    -- index (ignored when xColorType = 2)
                           , '0000'  -- nTintAndShade (0 = no change)
                           , utl_raw.substr(rColorCode, 2)     -- bRed, bGreen, bBlue
                           , utl_raw.substr(rColorCode, 1, 1)  -- bAlpha
                           );
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
    --recordTypeLabelMap(raw2int(RT_STRING)) := 'String';
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
    bitmask  raw(1) := BITMASKTABLE(bitNum);
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
    -- current record start
    stream.rstart := stream.offset;
    debug('RECORD INFO ['||to_char(stream.rstart,'FM0XXXXXXX')||']['||lpad(stream.rsize,6)||'] '||stream.rt);
  end;

  procedure seek_first (
    stream       in out nocopy Stream_T
  , record_type  in raw  
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
          -- flush
          dbms_lob.writeappend(stream.content, stream.buf_sz, stream.buf);
          stream.buf := bytes;
          stream.buf_sz := len;
        end if;
      end if;
    end;
  begin
    put(rt);
    put(rsize);
    if rec.is_lob then
      dbms_lob.copy(stream.content, rec.content, dbms_lob.getlength(rec.content), dbms_lob.getlength(stream.content) + 1);
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

  procedure write_XLString (
    rec       in out nocopy Record_T
  , strValue  in varchar2
  , stType    in pls_integer default ST_SIMPLE
  )
  is
    --amt     pls_integer;
    --cbuf    varchar2(32764);
    cch     pls_integer;
    csname  varchar2(30) := 'AL16UTF16LE';
    --offset  integer := 1;
    --rem     integer;
  begin
    if stType = ST_RICHSTR then
      write_record(rec, '00'); -- fRichStr, fExtStr, unused1
    end if;
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
    /*
    if str.is_lob then
      cch := least(dbms_lob.getlength(str.lobValue), 32767);
      write_record(rec, int2raw(cch)); -- cchCharacters
      rem := cch;
      -- rgchData
      while rem != 0 loop
        amt := least(8191, rem);
        dbms_lob.read(str.lobValue, amt, offset, cbuf);
        offset := offset + amt;
        rem := rem - amt;
        write_record(rec, utl_i18n.string_to_raw(cbuf, csname));
      end loop;
    else
      cch := length(str.strValue);
      write_record(rec, int2raw(cch)); -- cchCharacters
      -- rgchData
      if cch < 16384 then
        write_record(rec, utl_i18n.string_to_raw(str.strValue, csname));
      else
        write_record(rec, utl_i18n.string_to_raw(substr(str.strValue, 1, 16383), csname));
        write_record(rec, utl_i18n.string_to_raw(substr(str.strValue, 16384, 32766), csname));
        if cch = 32767 then
          write_record(rec, utl_i18n.string_to_raw(substr(str.strValue, 32767, 1), csname));
        end if;
      end if;
    end if;
    */
  end;

  procedure write_XLStringLob (
    rec       in out nocopy Record_T
  , lobValue  in clob
  , stType    in pls_integer default ST_SIMPLE
  )
  is
    amt     pls_integer;
    cbuf    varchar2(32764);
    cch     pls_integer;
    csname  varchar2(30) := 'AL16UTF16LE';
    offset  integer := 1;
    rem     integer;
  begin
    if stType = ST_RICHSTR then
      write_record(rec, '00');  -- fRichStr, fExtStr, unused1
    end if;    

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
  end;
  
  procedure write_RfX (
    rec       in out nocopy Record_T
  , firstRow  in pls_integer
  , firstCol  in pls_integer
  , lastRow   in pls_integer
  , lastCol   in pls_integer 
  )
  is
  begin
    write_record(rec, int2raw(firstRow));  -- rwFirst
    write_record(rec, int2raw(lastRow));   -- rwLast
    write_record(rec, int2raw(firstCol));  -- colFirst
    write_record(rec, int2raw(lastCol));   -- colLast    
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
    write_XLString(numFmt, fmt);           -- stFmtCode
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
    write_record(rec, int2raw( case when font.i then 2 else 0 end  -- bit1 : fItalic
                               , 2 ));
    write_record(rec, case when font.b then 'BC02' else '9001' end);  -- bls
    write_record(rec, '0000');  -- sss : None
    write_record(rec, int2raw(ExcelTypes.getUnderlineStyleId(nvl(font.u,'none')), 1));  -- uls
    write_record(rec, '00');    -- bFamily : Not applicable
    write_record(rec, '01');    -- bCharset : DEFAULT_CHARSET
    write_record(rec, '00');    -- unused
    write_record(rec, make_BrtColor(font.color));
    write_record(rec, '00');    -- bFontScheme : None
    write_XLString(rec, nvl(font.name, ExcelTypes.DEFAULT_FONT_FAMILY));  -- name
  
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
    xfId        in pls_integer default null
  , numFmtId    in pls_integer default 0
  , fontId      in pls_integer default 0
  , fillId      in pls_integer default 0
  , borderId    in pls_integer default 0
  , hAlignment  in varchar2 default null
  , vAlignment  in varchar2 default null
  , wrapText    in boolean default false
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
    write_record(xf, '00');  -- trot
    write_record(xf, '00');  -- indent
        
    write_record(xf, 
      bitVector( b0 => ExcelTypes.getHorizontalAlignmentId(nvl(hAlignment,'general'))  -- alc (3 bits)
               , b3 => ExcelTypes.getVerticalAlignmentId(nvl(vAlignment,'bottom'))     -- alcv (3 bits)
               , b6 => case when wrapText then 1 else 0 end  -- fWrap
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
      , case when hAlignment is not null or vAlignment is not null then 1 else 0 end
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
    write_XLString(rec, styleName);           -- stName
    return rec;
  end;
  
  function make_BundleSh (
    sheetId    in pls_integer
  , relId      in varchar2
  , sheetName  in varchar2
  )
  return Record_T
  is
    sh  Record_T := new_record(BRT_BUNDLESH);
  begin
    write_record(sh, '00000000');        -- hsState : VISIBLE
    write_record(sh, int2raw(sheetId));  -- iTabID
    write_XLString(sh, relId);           -- strRelID
    write_XLString(sh, sheetName);       -- strName
    return sh;
  end;
  
  function make_ExternSheet (
    links  in SupportingLinks_T
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_EXTERNSHEET);
  begin
    write_record(rec, int2raw(links.count));  -- cXti
    for i in links.first .. links.last loop
      -- Xti : 
      write_record(rec, int2raw(links(i).externalLink));  -- externalLink
      write_record(rec, int2raw(links(i).firstSheet));    -- firstSheet
      write_record(rec, int2raw(links(i).lastSheet));     -- lastSheet
    end loop;
    return rec;
  end;
  
  function make_FilterDatabase (
    bundleShIndex  in pls_integer
  , xti            in pls_integer
  , firstRow       in pls_integer
  , firstCol       in pls_integer
  , lastRow        in pls_integer
  , lastCol        in pls_integer
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_NAME);
  begin
    write_record(rec, 
      bitVector( 1  -- fHidden
               , 0  -- fFunc
               , 0  -- fOB
               , 0  -- fProc
               , 0  -- fCalcExp
               , 1  -- fBuiltin
               )
    );
    write_record(rec, '000000'); -- fgrp, fPublished, fWorkbookParam, fFutureFunction, reserved : all bits set to 0
    
    write_record(rec, '00'); -- chKey
    write_record(rec, int2raw(bundleShIndex)); -- itab
    write_XLString(rec, '_FilterDatabase'); -- name
    -- <BrtName.formula 
    write_record(rec, '0F000000'); -- cce
     
    write_record(rec, '3B');            -- PtgArea3d
    write_record(rec, int2raw(xti,2));  -- ixti
    -- <RgceAreaRel
    write_record(rec, int2raw(firstRow));    -- rowFirst
    write_record(rec, int2raw(lastRow));     -- rowLast
    write_record(rec, int2raw(firstcol,2));  -- columnFirst
    write_record(rec, int2raw(lastCol,2));   -- columnLast
    -- RgceAreaRel>
    
    write_record(rec, '00000000'); -- NameParsedFormula.cb
    -- BrtName.formula>
    
    write_record(rec, 'FFFFFFFF'); -- BrtName.comment : NULL
    
    -- no unusedstring1, description, helpTopic and unusedstring2 since fProc = 0
  
    return rec;
  end;
  
  function make_CalcProp (
    calcId  in pls_integer
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
      , 1 -- B: fRefA1
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
    write_RfX(rec, firstRow, firstCol, lastRow, lastCol);
    return rec;
  end;
  
  function make_ListPart (
    relId  in varchar2
  )
  return Record_T
  is
    rec  Record_T := new_record(BRT_LISTPART);
  begin
    write_XLString(rec, relId);
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
    write_RfX(rec, firstRow, firstCol, lastRow, lastCol);  -- rfxList
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
    write_XLString(rec, name);         -- stName
    write_XLString(rec, displayName);  -- stDisplayName
    write_XLString(rec, null);         -- stComment
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
    write_XLString(rec, fieldName);  -- stCaption
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
    -- stStyleName :
    if tableStyleName is not null then
      write_XLString(rec, tableStyleName);
    else
      write_record(rec, 'FFFFFFFF'); -- NULL
    end if;
    
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
  
  function add_SupportingLink (
    links         in out nocopy SupportingLinks_T
  , externalLink  in pls_integer
  , firstSheet    in pls_integer
  , lastSheet     in pls_integer default null
  )
  return pls_integer
  is
    idx  pls_integer := nvl(links.last, -1) + 1;
  begin
    links(idx).externalLink := externalLink;
    links(idx).firstSheet := firstSheet;
    links(idx).lastSheet := nvl(lastSheet, firstSheet);
    return idx;
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
    if lobValue is not null then
      write_XLStringLob(rec, lobValue);
    else
      write_XLString(rec, strValue);
    end if;
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
    stream  in out nocopy stream_t
  , str     in varchar2
  )
  is
    rec  Record_T := new_record(BRT_SSTITEM);
  begin
    write_XLString(rec, str, ST_RICHSTR);
    put_record(stream, rec);
  end;
  
  procedure put_defaultBookViews (
    stream  in out nocopy stream_t
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
    write_record(rec, '00000000');  -- itabFirst
    write_record(rec, '00000000');  -- itabCur
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
  )
  is
  begin
    put_record(stream, make_BundleSh(sheetId, relId, sheetName));
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
  
  procedure put_ExternSheet (
    stream  in out nocopy stream_t
  , links   in SupportingLinks_T
  )
  is
  begin
    put_record(stream, make_ExternSheet(links));
  end;

  procedure put_FilterDatabase (
    stream         in out nocopy stream_t
  , bundleShIndex  in pls_integer
  , xti            in pls_integer
  , firstRow       in pls_integer
  , firstCol       in pls_integer
  , lastRow        in pls_integer
  , lastCol        in pls_integer
  )
  is
  begin
    put_record(stream
             , make_FilterDatabase (
                 bundleShIndex
               , xti
               , firstRow
               , firstCol
               , lastRow
               , lastCol
               ));
  end;
  
  procedure put_CalcProp (
    stream  in out nocopy stream_t
  , calcId  in pls_integer
  )
  is
  begin
    put_record(stream, make_CalcProp(calcId));
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
  , hAlignment  in varchar2 default null
  , vAlignment  in varchar2 default null
  , wrapText    in boolean default false
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
               , hAlignment
               , vAlignment
               , wrapText
               ));    
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
    rk.fX100 := is_bit_set(utl_raw.substr(rk.RkNumber,1,1), 1);
    rk.fInt := is_bit_set(utl_raw.substr(rk.RkNumber,1,1), 2);
    
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
      str.fRichStr := is_bit_set(raw1, 1);
      str.fExtStr := is_bit_set(raw1, 2);
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
        , rpad(recordTypeLabelMap(stream.rt), 34)
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
    rcnt      pls_integer := 1;
    rw        pls_integer := ctx.curr_rw;
    col       pls_integer;
    num       number;
    str       String_T;
    emptyRow  boolean := true;
    
    cells  ExcelTableCellList := ExcelTableCellList();
    
    procedure read_col is
    begin
      col := read_int32(ctx.stream);
      skip(ctx.stream, 4); -- iStyleRef, fPhShow, reserved
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
          add_cell(anydata.ConvertNumber(num));
        
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
          add_cell(anydata.ConvertNumber(num));
          
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
          add_cell(anydata.ConvertNumber(num));
        
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
  
  function new_context (
    p_sst_part  in blob
  , p_cols      in varchar2 default null
  , p_firstRow  in pls_integer default null
  , p_lastRow   in pls_integer default null
  )
  return pls_integer
  is
    ctx     Context_T;
    ctx_id  pls_integer; 
  begin
      
    if p_sst_part is not null then
      read_sst(p_sst_part, ctx.sst);
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

begin
  
  init();

end xutl_xlsb;
/
