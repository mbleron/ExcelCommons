create or replace package body ExcelTypes is

  NAMED_COLORS          constant varchar2(4000) := 
  'aliceblue:F0F8FF;antiquewhite:FAEBD7;aqua:00FFFF;aquamarine:7FFFD4;azure:F0FFFF;beige:F5F5DC;bisque:FFE4C4;black:000000;blanchedalmond:FFEBCD;' ||
  'blue:0000FF;blueviolet:8A2BE2;brown:A52A2A;burlywood:DEB887;cadetblue:5F9EA0;chartreuse:7FFF00;chocolate:D2691E;coral:FF7F50;cornflowerblue:6495ED;' ||
  'cornsilk:FFF8DC;crimson:DC143C;cyan:00FFFF;darkblue:00008B;darkcyan:008B8B;darkgoldenrod:B8860B;darkgray:A9A9A9;darkgreen:006400;darkgrey:A9A9A9;' ||
  'darkkhaki:BDB76B;darkmagenta:8B008B;darkolivegreen:556B2F;darkorange:FF8C00;darkorchid:9932CC;darkred:8B0000;darksalmon:E9967A;darkseagreen:8FBC8F;' ||
  'darkslateblue:483D8B;darkslategray:2F4F4F;darkslategrey:2F4F4F;darkturquoise:00CED1;darkviolet:9400D3;deeppink:FF1493;deepskyblue:00BFFF;dimgray:696969;' ||
  'dimgrey:696969;dodgerblue:1E90FF;firebrick:B22222;floralwhite:FFFAF0;forestgreen:228B22;fuchsia:FF00FF;gainsboro:DCDCDC;ghostwhite:F8F8FF;gold:FFD700;' ||
  'goldenrod:DAA520;gray:808080;green:008000;greenyellow:ADFF2F;grey:808080;honeydew:F0FFF0;hotpink:FF69B4;indianred:CD5C5C;indigo:4B0082;ivory:FFFFF0;' ||
  'khaki:F0E68C;lavender:E6E6FA;lavenderblush:FFF0F5;lawngreen:7CFC00;lemonchiffon:FFFACD;lightblue:ADD8E6;lightcoral:F08080;lightcyan:E0FFFF;' ||
  'lightgoldenrodyellow:FAFAD2;lightgray:D3D3D3;lightgreen:90EE90;lightgrey:D3D3D3;lightpink:FFB6C1;lightsalmon:FFA07A;lightseagreen:20B2AA;lightskyblue:87CEFA;' ||
  'lightslategray:778899;lightslategrey:778899;lightsteelblue:B0C4DE;lightyellow:FFFFE0;lime:00FF00;limegreen:32CD32;linen:FAF0E6;magenta:FF00FF;maroon:800000;' ||
  'mediumaquamarine:66CDAA;mediumblue:0000CD;mediumorchid:BA55D3;mediumpurple:9370DB;mediumseagreen:3CB371;mediumslateblue:7B68EE;mediumspringgreen:00FA9A;' ||
  'mediumturquoise:48D1CC;mediumvioletred:C71585;midnightblue:191970;mintcream:F5FFFA;mistyrose:FFE4E1;moccasin:FFE4B5;navajowhite:FFDEAD;navy:000080;' ||
  'oldlace:FDF5E6;olive:808000;olivedrab:6B8E23;orange:FFA500;orangered:FF4500;orchid:DA70D6;palegoldenrod:EEE8AA;palegreen:98FB98;paleturquoise:AFEEEE;' ||
  'palevioletred:DB7093;papayawhip:FFEFD5;peachpuff:FFDAB9;peru:CD853F;pink:FFC0CB;plum:DDA0DD;powderblue:B0E0E6;purple:800080;rebeccapurple:663399;' ||
  'red:FF0000;rosybrown:BC8F8F;royalblue:4169E1;saddlebrown:8B4513;salmon:FA8072;sandybrown:F4A460;seagreen:2E8B57;seashell:FFF5EE;sienna:A0522D;silver:C0C0C0;' ||
  'skyblue:87CEEB;slateblue:6A5ACD;slategray:708090;slategrey:708090;snow:FFFAFA;springgreen:00FF7F;tan:D2B48C;teal:008080;thistle:D8BFD8;tomato:FF6347;' ||
  'turquoise:40E0D0;violet:EE82EE;wheat:F5DEB3;white:FFFFFF;steelblue:4682B4;whitesmoke:F5F5F5;yellow:FFFF00;yellowgreen:9ACD32';
  
  --DEFAULT_FONT_SIZE_PT  constant varchar2(256) := to_char(DEFAULT_FONT_SIZE, 'TM9', 'nls_numeric_characters=''. ''')||'pt';

  type simpleTypeMap_t is table of pls_integer index by varchar2(16);

  -- BEGIN CSS parser constants & structures

  CHAR_NL        constant varchar2(1) := chr(10);
  CHAR_CR        constant varchar2(1) := chr(13);
  CHAR_FF        constant varchar2(1) := chr(12);
  SURROGATE_MIN  constant pls_integer := to_number('D800','XXXX'); 
  SURROGATE_MAX  constant pls_integer := to_number('DFFF','XXXX');
  UNICODE_MAX    constant pls_integer := to_number('10FFFF','XXXXXX');
    
  T_EOF          constant pls_integer := -1;
  T_WHITESPACE   constant pls_integer := 0;
  T_STRING       constant pls_integer := 1;
  T_IDENT        constant pls_integer := 2;
  T_DIMENSION    constant pls_integer := 3;
  T_PERCENTAGE   constant pls_integer := 4;
  T_NUMBER       constant pls_integer := 5;
  T_HASH         constant pls_integer := 6;
  T_DELIM        constant pls_integer := 7;
  T_COMMA        constant pls_integer := 8;
  T_COLON        constant pls_integer := 9;
  T_SEMICOLON    constant pls_integer := 10;
  T_LEFT         constant pls_integer := 11;
  T_RIGHT        constant pls_integer := 12;
  T_FUNCTION     constant pls_integer := 13;
  
  CSS_PROP_VALUES_MIN        constant varchar2(256) := '''%s'' property expects at least %d value(s)';
  CSS_PROP_VALUES_MAX        constant varchar2(256) := '''%s'' property expects at most %d value(s)';
  CSS_INVALID_VALUE          constant varchar2(256) := 'Unsupported or invalid property value: ''%s''';
  CSS_UNEXPECTED_TOKEN_TYPE  constant varchar2(256) := 'Unexpected <%s> at position (%d,%d)';
  CSS_UNEXPECTED_TOKEN       constant varchar2(256) := 'Unexpected <%s>: ''%s'' at position (%d,%d)';
  CSS_BAD_RGB_COLOR          constant varchar2(256) := 'Invalid RGB color code: %s';
  
  type cssTokenLabelMap_t is table of varchar2(256) index by pls_integer;
  cssTokenLabels  cssTokenLabelMap_t;
  
  type pos_t is record (
    ln       pls_integer -- line number
  , cn       pls_integer -- column number
  , last_cn  pls_integer -- last column number of previous line
  );
    
  type token_t is record (
    t     pls_integer
  , v     varchar2(256)
  , nv    number
  , u     varchar2(256)
  --, args  serializedTokenList_t
  , argListId  pls_integer
  , pos   pos_t
  );
  
  type tokenList_t is table of token_t;
  type tokenListMap_t is table of tokenList_t index by pls_integer;
  type component_t is record (token token_t, tokenList tokenList_t, isList boolean := false);
  type componentList_t is table of component_t;
  type declaration_t is record (name varchar2(256), v componentList_t, argListMap tokenListMap_t);
  type declarationList_t is table of declaration_t;
  
  type rgbColor_t is record (r pls_integer, g pls_integer, b pls_integer, a number);
  type cssBorderSide_t is record (style varchar2(256) := 'none', width varchar2(256) := 'medium', color varchar2(256));
  type cssBorder_t is record (top cssBorderSide_t, right cssBorderSide_t, bottom cssBorderSide_t, left cssBorderSide_t);
  type cssFont_t is record (family varchar2(256)/* := DEFAULT_FONT_FAMILY*/, sz varchar2(256)/* := DEFAULT_FONT_SIZE_PT*/, style varchar2(256) := 'normal', weight varchar2(256) := 'normal');
  type cssTextDecoration_t is record (line varchar2(256) := 'none', style varchar2(256) := 'solid');
  type cssMsoPattern_t is record (patternType varchar2(256) := 'none', color varchar2(256));
  
  type colorStop_t is record (colorHint number, color varchar2(256), pct1 number, pct2 number);
  type colorStopList_t is table of colorStop_t;
  type linearGradient_t is record (angle number, colorStopList colorStopList_t);
  
  type cssStyle_t is record (
    border           cssBorder_t
  , font             cssFont_t
  , textDecoration   cssTextDecoration_t
  , verticalAlign    varchar2(256)
  , textAlign        varchar2(256)
  , whiteSpace       varchar2(256) := 'pre'
  , msoPattern       cssMsoPattern_t
  , msoNumberFormat  varchar2(256)
  , color            varchar2(256)
  , backgroundColor  varchar2(256)
  , backgroundImage  linearGradient_t
  );
  
  functionArgListCache  tokenListMap_t;
  
  -- END CSS parser constants & structures
  
  colorMap           colorMap_t;
  underlineStyleMap  simpleTypeMap_t;
  fillPatternMap     simpleTypeMap_t;
  borderStyleMap     simpleTypeMap_t;
  hAlignmentMap      simpleTypeMap_t;
  vAlignmentMap      simpleTypeMap_t;
  
  debug_enabled  boolean := false;
  
  procedure setDebug (p_status in boolean)
  is
  begin
    debug_enabled := nvl(p_status, false);
  end;

  procedure debug (message in varchar2)
  is
  begin
    if debug_enabled then
      dbms_output.put_line(message);
    end if;
  end;

  procedure loadCssTokenLabels is
  begin
    cssTokenLabels(T_EOF) := 'EOF-token';
    cssTokenLabels(T_WHITESPACE) := 'whitespace-token';
    cssTokenLabels(T_STRING) := 'string-token';
    cssTokenLabels(T_IDENT) := 'ident-token';
    cssTokenLabels(T_DIMENSION) := 'dimension-token';
    cssTokenLabels(T_PERCENTAGE) := 'percentage-token';
    cssTokenLabels(T_NUMBER) := 'number-token';
    cssTokenLabels(T_HASH) := 'hash-token';
    cssTokenLabels(T_DELIM) := 'delim-token';
    cssTokenLabels(T_COMMA) := 'comma-token';
    cssTokenLabels(T_COLON) := 'colon-token';
    cssTokenLabels(T_SEMICOLON) := 'semicolon-token';
    cssTokenLabels(T_LEFT) := 'left-parenthesis-token';
    cssTokenLabels(T_RIGHT) := 'right-parenthesis-token';
    cssTokenLabels(T_FUNCTION) := 'function-token';
  end;
  
  procedure initColorMap 
  is
    token  varchar2(32);
    p1     pls_integer := 1;
    p2     pls_integer;  
    i      pls_integer;
  begin
    loop
      p2 := instr(NAMED_COLORS, ';', p1);
      if p2 = 0 then
        token := substr(NAMED_COLORS, p1);
      else
        token := substr(NAMED_COLORS, p1, p2-p1);    
        p1 := p2 + 1;
      end if;
      i := instr(token,':');
      colorMap(substr(token,1,i-1)) := substr(token,i+1);
      exit when p2 = 0;
    end loop;   
  end;
  
  procedure initialize
  is
  begin
    
    -- CSS token labels
    loadCssTokenLabels;
    
    -- named colors
    initColorMap;
    
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

  procedure error (
    msg in varchar2
  , arg1 in varchar2 default null
  , arg2 in varchar2 default null
  , arg3 in varchar2 default null
  , arg4 in varchar2 default null
  ) is
  begin
    raise_application_error(-20000, utl_lms.format_message(msg, arg1, arg2, arg3, arg4));
  end;

  function isValidColorCode (p_colorCode in varchar2) return boolean is
  begin
    return regexp_like(upper(p_colorCode), '^[0-9A-F]{6}([0-9A-F]{2})?$');
  end;
  
  function isValidColorName (p_colorName in varchar2) return boolean is
  begin
    return colorMap.exists(lower(p_colorName));
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

  function getColorMap return colorMap_t is
  begin
    return colorMap;
  end;

  function getColorCode (p_colorName in varchar2) return varchar2 is
  begin
    return colorMap(lower(p_colorName));
  end;
  
  function makeRgbColor (r in uint8, g in uint8, b in uint8, a in number default null) return varchar2 is
  begin
    if a not between 0 and 1 then
      error('Alpha component must be a number between 0 and 1.');
    end if;
    return to_char(r * 65536 + g * 256 + b, 'FM0XXXXX')||to_char(a*255, 'FM0X');
  end;

  function makeRgbColor (color in rgbColor_t) return varchar2 is
  begin
    return makeRgbColor(color.r, color.g, color.b, color.a);
  end;

  -- blend optional alpha channel with opaque white background
  procedure blendRgbAlpha (color in out nocopy rgbColor_t) is
  begin
    if color.a < 1 then
      color.r := (1 - color.a)*255 + color.a * color.r;
      color.g := (1 - color.a)*255 + color.a * color.g;
      color.b := (1 - color.a)*255 + color.a * color.b;
    end if;
    color.a := null;
  end;
  
  function parseRgbColor (rgbCode in varchar2, blendAlpha in boolean default false) return rgbColor_t is
    color  rgbColor_t;
  begin
    color.r := to_number(substr(rgbCode,1,2),'XX');
    color.g := to_number(substr(rgbCode,3,2),'XX');
    color.b := to_number(substr(rgbCode,5,2),'XX');
    color.a := to_number(substr(rgbCode,7,2),'XX')/255;
    if blendAlpha then
      blendRgbAlpha(color);
    end if;
    return color;
  end;

  -- mix two RGBA colors
  function mixRgbColor (rgb1 in varchar2, rgb2 in varchar2) return varchar2 is
    c1  rgbColor_t := parseRgbColor(rgb1, true);
    c2  rgbColor_t := parseRgbColor(rgb2, true);
  begin
    return makeRgbColor((c1.r + c2.r)/2, (c1.g + c2.g)/2, (c1.b + c2.b)/2);
  end;

  function validateColor (
    colorSpec  in varchar2
  )
  return varchar2
  is
    rgbCode  varchar2(8);
  begin
    if colorSpec is not null then
      -- RGB color code?
      if substr(colorSpec,1,1) = '#' then
        rgbCode := upper(substr(colorSpec,2)); 
        if rgbCode is null or not isValidColorCode(rgbCode) then
          error(CSS_BAD_RGB_COLOR, colorSpec);
        end if;
        -- blend optional alpha channel with white background
        if length(rgbCode) = 8 then
          rgbCode := makeRgbColor(parseRgbColor(rgbCode, true));
        end if;
        -- opaque by default
        rgbCode := 'FF' || rgbCode;
      elsif isValidColorName(colorSpec) then
        rgbCode := 'FF' || getColorCode(lower(colorSpec));
      elsif regexp_like(colorSpec,'^theme:\d+$') then
        rgbCode := colorSpec;
      else
        error('Invalid color code: %s', colorSpec);
      end if;
    end if;
    return rgbCode;
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
  
  procedure unexpectedToken (token in token_t) is
  begin
    if token.v is not null then
      error(CSS_UNEXPECTED_TOKEN, cssTokenLabels(token.t), token.v, token.pos.ln, token.pos.cn);
    else
      error(CSS_UNEXPECTED_TOKEN_TYPE, cssTokenLabels(token.t), token.pos.ln, token.pos.cn);
    end if;
  end;

  procedure assertNotIsList (comp in component_t) is
  begin
    if comp.isList then
      error('Unsupported component type: value list');
    end if;
  end;

  procedure assertValueCount (
    propName    in varchar2
  , valueCount  in pls_integer
  , valueMax    in pls_integer
  , valueMin    in pls_integer default null
  )
  is
  begin
    if valueCount < valueMin then
      error(CSS_PROP_VALUES_MIN, propName, valueMin);
    elsif valueCount > valueMax then
      error(CSS_PROP_VALUES_MAX, propName, valueMax);
    end if;    
  end;

  procedure stringWrite (
    buf  in out nocopy varchar2
  , str  in varchar2
  )
  is
  begin
    buf := buf || str;
  end;

  procedure setBorderContent (
    border  in out nocopy CT_Border
  )
  is
    function getBorderPrContent (borderName in varchar2, borderPr in CT_BorderPr)
    return varchar2
    is
    begin
      return '<' || borderName || 
             case when nvl(borderPr.style, 'none') != 'none' then ' style="'||borderPr.style||'"' end || 
             case when borderPr.color is not null then '><color rgb="'||borderPr.color||'"/></'||borderName||'>' else '/>' end;
    end;
  begin
    border.content := null;
    stringWrite(border.content, '<border>');
    stringWrite(border.content, getBorderPrContent('left', border.left));
    stringWrite(border.content, getBorderPrContent('right', border.right));
    stringWrite(border.content, getBorderPrContent('top', border.top));
    stringWrite(border.content, getBorderPrContent('bottom', border.bottom));
    stringWrite(border.content, '</border>');    
  end;

  procedure setFontContent (
    font  in out nocopy CT_Font
  )
  is
  begin
    font.content := null;
    stringWrite(font.content, '<font>');
    stringWrite(font.content, '<sz val="'||to_char(nvl(font.sz, DEFAULT_FONT_SIZE))||'"/>');
    stringWrite(font.content, '<name val="'||nvl(font.name, DEFAULT_FONT_FAMILY)||'"/>');
    if font.b then
      stringWrite(font.content, '<b/>');
    end if;
    if font.i then
      stringWrite(font.content, '<i/>');
    end if;
    if font.color is not null then
      if font.color like 'theme:%' then
        stringWrite(font.content, '<color theme="'||regexp_substr(font.color,'\d+$')||'"/>');
      else
        stringWrite(font.content, '<color rgb="'||font.color||'"/>');
      end if;
    end if;
    stringWrite(font.content, '<u val="'||nvl(font.u, 'none')||'"/>');
    stringWrite(font.content, '</font>');    
  end;

  procedure setFillContent (
    fill  in out nocopy CT_Fill
  )
  is
  begin
    fill.content := null;
    case fill.fillType
    when FT_PATTERN then
      
      stringWrite(fill.content, '<fill><patternFill patternType="'||fill.patternFill.patternType||'">');
      if fill.patternFill.fgColor is not null then
        stringWrite(fill.content, '<fgColor rgb="'||fill.patternFill.fgColor||'"/>');
      end if;
      if fill.patternFill.bgColor is not null then
        stringWrite(fill.content, '<bgColor rgb="'||fill.patternFill.bgColor||'"/>');
      end if;
      stringWrite(fill.content, '</patternFill></fill>');
      
    when FT_GRADIENT then
      
      stringWrite(fill.content, '<fill><gradientFill degree="'||to_char(nvl(fill.gradientFill.degree,0), 'TM9', 'nls_numeric_characters=''. ''')||'">');
      for i in 1 .. fill.gradientFill.stops.count loop
        stringWrite(fill.content, '<stop position="'||to_char(fill.gradientFill.stops(i).position, 'TM9', 'nls_numeric_characters=''. ''')||'">');
        stringWrite(fill.content, '<color rgb="'||fill.gradientFill.stops(i).color||'"/>');
        stringWrite(fill.content, '</stop>');
      end loop;
      stringWrite(fill.content, '</gradientFill></fill>');      
    
    else
      error('Invalid fill type.');
    end case;
  end;

  procedure setAlignmentContent (
    alignment  in out nocopy CT_CellAlignment
  )
  is
  begin
    alignment.content := null;
    if coalesce(alignment.horizontal, alignment.vertical) is not null or alignment.wrapText then
      stringWrite(alignment.content, '<alignment');
      if alignment.horizontal is not null then
        stringWrite(alignment.content, ' horizontal="'||alignment.horizontal||'"');
      end if;
      if alignment.vertical is not null then
        stringWrite(alignment.content, ' vertical="'||alignment.vertical||'"');
      end if;
      if alignment.wrapText then
        stringWrite(alignment.content, ' wrapText="1"');
      end if;
      stringWrite(alignment.content, '/>');
    end if;    
  end;

  function makeBorderPr (
    p_style  in varchar2 default null
  , p_color  in varchar2 default null
  )
  return CT_BorderPr
  is
    borderPr  CT_BorderPr;
  begin
    borderPr.style := p_style;
    borderPr.color := validateColor(p_color);
    return borderPr;
  end;
  
  function makeBorder (
    p_left    in CT_BorderPr default makeBorderPr()
  , p_right   in CT_BorderPr default makeBorderPr()
  , p_top     in CT_BorderPr default makeBorderPr()
  , p_bottom  in CT_BorderPr default makeBorderPr()
  )
  return CT_Border
  is
    border  CT_Border;
  begin
    border.left := p_left;
    border.right := p_right;
    border.top := p_top;
    border.bottom := p_bottom;
    setBorderContent(border);
    return border;
  end;
  
  function makeBorder (
    p_style  in varchar2
  , p_color  in varchar2 default null
  )
  return CT_Border
  is
    borderPr  CT_BorderPr := makeBorderPr(p_style, p_color);
  begin
    return makeBorder(borderPr, borderPr, borderPr, borderPr);
  end;

  function makeFont (
    p_name   in varchar2 default null
  , p_sz     in pls_integer default null
  , p_b      in boolean default false
  , p_i      in boolean default false
  , p_color  in varchar2 default null
  , p_u      in varchar2 default null
  )
  return CT_Font
  is
    font  CT_Font;
  begin
    if p_name is not null then
      font.name := p_name;
    end if;
    if p_sz is not null then
      font.sz := p_sz;
    end if;
    font.b := nvl(p_b, false);
    font.i := nvl(p_i, false);
    
    if p_u is not null then 
      if isValidUnderlineStyle(p_u) then
        font.u := p_u;
      else
        error('Invalid underline style: %s', p_u);
      end if;
    end if;
    
    font.color := validateColor(p_color);
    setFontContent(font);
    return font;
  end;

  function makePatternFill (
    p_patternType  in varchar2
  , p_fgColor      in varchar2 default null
  , p_bgColor      in varchar2 default null
  )
  return CT_Fill
  is
    fill  CT_Fill;
  begin
    fill.fillType := FT_PATTERN;
    fill.patternFill.patternType := nvl(p_patternType, 'none');
    fill.patternFill.fgColor := validateColor(p_fgColor);
    fill.patternFill.bgColor := validateColor(p_bgColor);
    setFillContent(fill);
    return fill;
  end;

  function makeGradientStop (
    p_position  in number
  , p_color     in varchar2
  )
  return CT_GradientStop
  is
    stop  CT_GradientStop;
  begin
    if p_position between 0 and 1 then
      stop.position := p_position;
    else
      error('Gradient stop position must be a number between 0 and 1.');
    end if;
    stop.color := validateColor(p_color);
    return stop;
  end;
  
  function makeGradientFill (
    p_degree  in number default null
  , p_stops   in CT_GradientStopList default null
  )
  return CT_Fill
  is
    fill  CT_Fill;
  begin
    fill.fillType := FT_GRADIENT;
    fill.gradientFill.degree := mod(nvl(p_degree, 0), 360);
    fill.gradientFill.stops := nvl(p_stops, CT_GradientStopList());
    setFillContent(fill);
    return fill;
  end;

  procedure addGradientStop (
    p_fill      in out nocopy CT_Fill
  , p_position  in number
  , p_color     in varchar2
  )
  is
  begin
    if p_fill.fillType != FT_GRADIENT then
      error('Invalid fill type');
    end if;
    p_fill.gradientFill.stops.extend;
    p_fill.gradientFill.stops(p_fill.gradientFill.stops.last) := makeGradientStop(p_position, p_color);
    setFillContent(p_fill);
  end;

  function makeAlignment (
    p_horizontal  in varchar2 default null
  , p_vertical    in varchar2 default null
  , p_wrapText    in boolean default false
  )
  return CT_CellAlignment
  is
    alignment  CT_CellAlignment;
  begin
    alignment.horizontal := p_horizontal;
    alignment.vertical := p_vertical;
    alignment.wrapText := p_wrapText;
    setAlignmentContent(alignment);
    return alignment;
  end;
  
  function mergeBorders (masterBorder in CT_Border, border in CT_Border) return CT_Border is
    mergedBorder  CT_Border := masterBorder;
    
    procedure mergeBorderPr (masterBorderPr in out nocopy CT_BorderPr, borderPr in CT_BorderPr) is
    begin
      if borderPr.style != 'none' then
        masterBorderPr.style := borderPr.style;
      end if;
      if borderPr.color is not null then
        masterBorderPr.color := borderPr.color;
      end if;
    end;
    
  begin
    mergeBorderPr(mergedBorder.left, border.left);
    mergeBorderPr(mergedBorder.right, border.right);
    mergeBorderPr(mergedBorder.top, border.top);
    mergeBorderPr(mergedBorder.bottom, border.bottom);
    setBorderContent(mergedBorder);
    return mergedBorder;
  end;
  
  function applyBorderSide (
    border  in CT_Border
  , top     in boolean
  , right   in boolean
  , bottom  in boolean
  , left    in boolean  
  )
  return CT_Border
  is
    newBorder  CT_Border;
  begin
    if top then
      newBorder.top := border.top;
    end if;
    if right then
      newBorder.right := border.right;
    end if;
    if bottom then
      newBorder.bottom := border.bottom;
    end if;
    if left then
      newBorder.left := border.left;
    end if;
    setBorderContent(newBorder);
    return newBorder;
  end;

  function mergeFonts (masterFont in CT_Font, font in CT_Font) return CT_Font is
    mergedFont  CT_Font := masterFont;
  begin  
    if font.name is not null then
      mergedFont.name := font.name;
    end if;
    if font.b then
      mergedFont.b := font.b;
    end if;
    if font.i then
      mergedFont.i := font.i;
    end if;
    if font.u != 'none' then
      mergedFont.u := font.u;
    end if;
    if font.color is not null then
      mergedFont.color := font.color;
    end if;
    if font.sz is not null then
      mergedFont.sz := font.sz;
    end if;
    setFontContent(mergedFont);
    return mergedFont;
  end;

  function mergePatternFills (masterFill in CT_Fill, fill in CT_Fill) return CT_Fill is
    mergedFill  CT_Fill := fill;
  begin
    -- if not set, apply patternType from master
    if mergedFill.patternFill.patternType = 'none' then
      mergedFill.patternFill.patternType := masterFill.patternFill.patternType;
    end if;
    
    -- if not set, apply fgColor from master
    -- reminder: in 'solid' pattern, fg and bg are reversed
    if mergedFill.patternFill.fgColor is null then
      if masterFill.patternFill.patternType = 'solid' and mergedFill.patternFill.patternType != 'solid' then
        mergedFill.patternFill.fgColor := masterFill.patternFill.bgColor;
      else
        mergedFill.patternFill.fgColor := masterFill.patternFill.fgColor;
      end if;
    end if;
    
    -- if not set, apply bgColor from master
    if mergedFill.patternFill.bgColor is null then
      if masterFill.patternFill.patternType = 'solid' and mergedFill.patternFill.patternType != 'solid' then
        mergedFill.patternFill.bgColor := masterFill.patternFill.fgColor;
      else
        mergedFill.patternFill.bgColor := masterFill.patternFill.bgColor;
      end if;
    end if;
    
    setFillContent(mergedFill);
    
    return mergedFill;

  end;

  function mergeAlignments (masterAlignment in CT_CellAlignment, alignment in CT_CellAlignment) return CT_CellAlignment is
    mergedAlignment  CT_CellAlignment := alignment;
  begin
    if mergedAlignment.horizontal is null then
      mergedAlignment.horizontal := masterAlignment.horizontal;
    end if;
    if mergedAlignment.vertical is null then
      mergedAlignment.vertical := masterAlignment.vertical;
    end if;
    
    if not mergedAlignment.wrapText then
      mergedAlignment.wrapText := masterAlignment.wrapText;
    end if;
    
    setAlignmentContent(mergedAlignment);
    return mergedAlignment;
  end;
  
  function parseRawCss (css in varchar2)
  return declarationList_t
  is
    
    input varchar2(32767) := css;
    
    type cstream_t is record (content varchar2(32767), p pls_integer := 1, c varchar2(1 char), pos pos_t);
    cstream  cstream_t;
    
    type tstream_t is record (tokens tokenList_t, idx pls_integer := 1, currentToken token_t);
    tstream  tstream_t;
    
    currentToken  token_t;
    comp          component_t;
    decl          declaration_t;
    decls         declarationList_t;
    
    inList        boolean;
    
    procedure append (buf in out nocopy varchar2, c in varchar2) is
    begin
      buf := buf || c;
    end;
    
    procedure appendToken (tlist in out nocopy tokenList_t, token in token_t) is
    begin
      tlist.extend;
      tlist(tlist.last) := token;
    end;

    procedure consumeNext is
    begin

      if cstream.c = CHAR_NL then
        cstream.pos.ln := cstream.pos.ln + 1;
        cstream.pos.last_cn := cstream.pos.cn;
        cstream.pos.cn := 0;
      end if;

      cstream.c := substr(cstream.content, cstream.p, 1);
      cstream.p := cstream.p + 1;
      
      cstream.pos.cn := cstream.pos.cn + 1;
      
    end;

    function consumeNext return varchar2 is
    begin
      consumeNext;
      return cstream.c;
    end;
    
    procedure reconsumeCurrent is
    begin
      cstream.p := cstream.p - 1;
      cstream.c := substr(cstream.content, cstream.p - 1, 1);
      
      cstream.pos.cn := cstream.pos.cn - 1;
      if cstream.pos.cn = 0 and cstream.pos.ln > 1 then
        cstream.pos.ln := cstream.pos.ln - 1;
        cstream.pos.cn := cstream.pos.last_cn;
      end if;
      
    end;
    
    function nextChar (pos in pls_integer default 1) return varchar2 is
    begin
      return substr(cstream.content, cstream.p + nvl(pos,1) - 1, 1);
    end;
    
    function newToken (
      t in varchar2
    , v in varchar2 default null
    , p in pos_t default null
    ) 
    return token_t is
      token  token_t;
    begin
      token.t := t;
      token.v := v;
      if p.ln is not null then
        token.pos := p;
      else
        token.pos := cstream.pos;
      end if;
      return token;
    end;
    
    procedure consumeNextToken is
    begin
      tstream.currentToken := tstream.tokens(tstream.idx);
      tstream.idx := tstream.idx + 1;
      debug(cssTokenLabels(tstream.currentToken.t)||'='||tstream.currentToken.v);
    end;

    function consumeNextToken return token_t is
    begin
      consumeNextToken;
      return tstream.currentToken;
    end;
    
    function nextToken return token_t is
    begin
      return tstream.tokens(tstream.idx);
    end;
    
    function codePointToChar (u in pls_integer) return varchar2 is
    begin
      if u < 65536 then -- U+010000
        return to_char(unistr('\'||to_char(u,'FMXXXX')));
      else
        -- convert code point to high and low UTF-16 surrogates
        return to_char(unistr(
                   '\' || to_char(trunc((u - 65536)/1024) + 55296, 'FMXXXX') ||
                   '\' || to_char(mod((u - 65536), 1024) + 56320, 'FMXXXX')
               ));
      end if;
    end;
    
    -- 4.2. Definitions : whitespace
    function isWhitespace (c in varchar2) return boolean is
    begin
      return ( c in (CHAR_NL, chr(9), chr(32)) );
    end;
    
    function isDigit (c in varchar2) return boolean is
    begin
      return ( c between '0' and '9' );
    end;
    
    function isHexDigit (c in varchar2) return boolean is
    begin
      return regexp_like(c, '[0-9A-Fa-f]');
    end;
    
    -- 4.2. Definitions : ident-start code point
    function isIdentStartCodePoint (c in varchar2) return boolean is
    begin
      return regexp_like(c, '[a-zA-Z_]') or nlssort(c, 'NLS_SORT=UNICODE_BINARY') >= hextoraw('0080');
    end;

    -- 4.2. Definitions : ident code point
    function isIdentCodePoint (c in varchar2) return boolean is
    begin
      return isIdentStartCodePoint(c) or regexp_like(c, '\d') or c = '-';
    end;
    
    -- 4.3.8. Check if two code points are a valid escape
    function isValidEscape (c1 varchar2, c2 varchar2) return boolean is
    begin
      if c1 = '\' then
        if c2 = CHAR_NL then
          return false;
        end if;
        return true;
      else
        return false;
      end if;
    end;
    
    -- 4.3.9. Check if three code points would start an ident sequence
    function startsIdentSequence (c1 varchar2, c2 varchar2, c3 varchar2) return boolean is
    begin
      case 
      when c1 = '-' then
        if isIdentStartCodePoint(c2) or c2 = '-' or isValidEscape(c2,c3) then
          return true;
        else
          return false;
        end if;
      when isIdentStartCodePoint(c1) then
        return true;
      when c1 = '\' then
        return isValidEscape(c1,c2);
      else
        return false;
      end case;
    end;
    
    -- 4.3.10. Check if three code points would start a number
    function startsNumber (c1 varchar2, c2 varchar2, c3 varchar2) return boolean is
    begin
      case 
      when c1 in ('+','-') then
        return ( isDigit(c2) or c2 = '.' and isDigit(c3) );
      when c1 = '.' then
        return isDigit(c2);
      when isDigit(c1) then
        return true;
      else
        return false;
      end case;
    end;
    
    -- 4.3.7. Consume an escaped code point
    function readEscaped return varchar2 is
      hexDigits  varchar2(6);
      n          pls_integer;
      c          varchar2(1 char);
    begin
      consumeNext;
      if isHexDigit(cstream.c) then
        append(hexDigits, cstream.c);
        while length(hexDigits) < 6 and isHexDigit(nextChar) loop
          consumeNext;
          append(hexDigits, cstream.c);
        end loop;
        if isWhitespace(nextChar) then
          consumeNext;
        end if;
        n := to_number(hexDigits, 'XXXXXX');
        if n = 0 or n between SURROGATE_MIN and SURROGATE_MAX or n > UNICODE_MAX then
          n := 65533; -- U+FFFD REPLACEMENT CHARACTER
        end if;
        c := codePointToChar(n);
              
      elsif cstream.c is null then
        unexpectedToken(newToken(T_EOF, p => cstream.pos));
        
      else
        c := cstream.c;
        
      end if;
      return c;
    end;
    
    -- 4.3.1. Consume a token - whitespace
    function readWhitespaceToken return token_t is
      strt  pos_t := cstream.pos;
    begin
      while isWhitespace(nextChar) loop
        consumeNext;
      end loop;
      return newToken(T_WHITESPACE, p => strt);
    end;

    -- 4.3.5. Consume a string token
    function readStringToken return token_t is
      token  token_t := newToken(T_STRING, p => cstream.pos);
      endingChar  varchar2(1 char) := cstream.c;
      c           varchar2(1 char);
    begin
      loop
        consumeNext;
        if cstream.c = endingChar then
          exit;
        elsif cstream.c is null then  
          unexpectedToken(newToken(T_EOF, p => cstream.pos));
        elsif cstream.c = CHAR_NL then
          error('Invalid code point found at position (%d,%d)', cstream.pos.ln, cstream.pos.cn);
        elsif cstream.c = '\' then
          c := nextChar;
          if c is null then
            null;
          elsif c = CHAR_NL then
            consumeNext;
          else
            c := readEscaped;
            append(token.v, c);
          end if;
        else
          append(token.v, cstream.c);
        end if;
      end loop;
      return token;
    end;
    
    -- 4.3.11. Consume an ident sequence
    function readIdentSequence return varchar2 is
      res  varchar2(256);
    begin
      loop
        consumeNext;
        if isIdentCodePoint(cstream.c) then
          append(res, cstream.c);
        elsif isValidEscape(cstream.c, nextChar) then
          append(res, readEscaped);
        else
          reconsumeCurrent;
          exit;
        end if;
      end loop;
      return res;
    end;
    
    -- 4.3.4. Consume an ident-like token
    function readIdentToken return token_t is
      token  token_t;
    begin
      token.pos.cn := cstream.pos.cn + 1;
      token.pos.ln := cstream.pos.ln;
      token.v := readIdentSequence;
      if nextChar = '(' then
        consumeNext;
        token.t := T_FUNCTION;
      else
        token.t := T_IDENT;
      end if;
      return token;
    end;
    
    -- 4.3.3. Consume a numeric token
    function readNumericToken return token_t is
      token  token_t;
    begin
      token.pos.cn := cstream.pos.cn + 1;
      token.pos.ln := cstream.pos.ln;
      
      if nextChar in ('+','-') then
        consumeNext;
        append(token.v, cstream.c);
      end if;
      while isDigit(nextChar) loop
        consumeNext;
        append(token.v, cstream.c);
      end loop;
      if nextChar = '.' and isDigit(nextChar(2)) then
        consumeNext;
        append(token.v, cstream.c);
        consumeNext;
        append(token.v, cstream.c);
        while isDigit(nextChar) loop
          consumeNext;
          append(token.v, cstream.c);
        end loop;      
      end if;
      if nextChar(1) in ('E','e') 
         and ( nextChar(2) in ('+','-') and isDigit(nextChar(3))
            or isDigit(nextChar(2)) )
      then
        consumeNext;
        append(token.v, cstream.c);
        if nextChar(1) in ('+','-') then
          consumeNext;
          append(token.v, cstream.c);        
        end if;
        consumeNext;
        append(token.v, cstream.c);
        while isDigit(nextChar) loop
          consumeNext;
          append(token.v, cstream.c);
        end loop;       
      end if;
      token.nv := to_number(token.v);
      
      if startsIdentSequence(nextChar(1), nextChar(2), nextChar(3)) then
        token.t := T_DIMENSION;
        token.u := readIdentSequence;
        append(token.v, token.u);
      elsif nextChar = '%' then
        consumeNext;
        append(token.v, cstream.c);
        token.t := T_PERCENTAGE;
      else
        token.t := T_NUMBER;
      end if;
      
      return token;
    end;
    
    -- 4.3.1. Consume a token
    procedure tokenize (cssContent in varchar2) is
      c       varchar2(1 char);
      token   token_t;
    begin
      
      cstream.pos.ln := 1;
      cstream.pos.cn := 0;
      cstream.content := cssContent;
    
      tstream.tokens := tokenList_t();
    
      loop
        
        c := consumeNext;
        
        case
        when c is null then
          token := newToken(T_EOF);
        when isWhitespace(c) then
          token := readWhitespaceToken;
        when c in ('"', '''') then
          token := readStringToken;
        when c = '#' then
          if isIdentCodePoint(nextChar) or isValidEscape(nextChar(1),nextChar(2)) then
            token := newToken(T_HASH, readIdentSequence);
          else
            token := newToken(T_DELIM, c);
          end if;
          
        when c = '(' then
          token := newToken(T_LEFT);
        
        when c = ')' then
          token := newToken(T_RIGHT);
        
        when c = '+' then
          if startsNumber(nextChar(1),nextChar(2),nextChar(3)) then
            reconsumeCurrent;
            token := readNumericToken;
          else
            token := newToken(T_DELIM, c);
          end if;
          
        when c = ',' then
          token := newToken(T_COMMA);
          
        when c = '-' then
          if startsNumber(nextChar(1),nextChar(2),nextChar(3)) then
            reconsumeCurrent;
            token := readNumericToken;
          elsif startsIdentSequence(nextChar(1),nextChar(2),nextChar(3)) then
            reconsumeCurrent;
            token := readIdentToken;
          else
            token := newToken(T_DELIM, c);
          end if;
          
        when c = '.' then
          if startsNumber(nextChar(1),nextChar(2),nextChar(3)) then
            reconsumeCurrent;
            token := readNumericToken;
          else
            token := newToken(T_DELIM, c);
          end if;
          
        when c = ':' then
          token := newToken(T_COLON); 

        when c = ';' then
          token := newToken(T_SEMICOLON);
          
        when c = '\' then
          if isValidEscape(c,nextChar(1)) then
            reconsumeCurrent;
            token := readIdentToken;
          else
            error('Invalid escape sequence found at position (%d,%d)', cstream.pos.ln, cstream.pos.cn);
          end if;
          
        when isDigit(c) then
          reconsumeCurrent;  
          token := readNumericToken;
          
        when isIdentStartCodePoint(c) then
          reconsumeCurrent;
          token := readIdentToken;
          
        else
          token := newToken(T_DELIM, c);
                
        end case;
        
        tstream.tokens.extend;
        tstream.tokens(tstream.tokens.last) := token;
        
        exit when token.t = T_EOF;
      
      end loop;
    
    end;
    
    -- 5.4.9. Consume a function
    procedure consumeFunction (funcToken in out nocopy token_t) is
      token  token_t;
      args   tokenList_t := tokenList_t();
    begin
      -- skip leading whitespace: 
      -- there should be at most one whitespace at this point since they've been collapsed earlier in tokenize() procedure
      if nextToken().t = T_WHITESPACE then
        consumeNextToken;
      end if;
      -- read values
      loop
        token := consumeNextToken;
        exit when token.t = T_RIGHT;
        if token.t = T_EOF then
          unexpectedToken(token);
        else
          --funcToken.args.extend;
          --funcToken.args(funcToken.args.last) := serializeToken(token);
          
          if token.t = T_FUNCTION then
            consumeFunction(token);
          end if;
          
          args.extend;
          args(args.last) := token;          
          
        end if;
      end loop;
      -- trim whitespace
      if args.count > 0 and args(args.last).t = T_WHITESPACE then
        args.trim;
      end if;
      -- push back right parenthesis token
      args.extend;
      args(args.last) := token;
      
      -- save arg list handle
      funcToken.argListId := nvl(decl.argListMap.last, 0) + 1;
      decl.argListMap(funcToken.argListId) := args;
      
      debug('======================================');
      debug('function: '||funcToken.v);
      debug('======================================');
      for i in 1 .. args.count loop
        debug(cssTokenLabels(args(i).t)||'='||args(i).v);
      end loop;
      debug('======================================');
      
    end;

  begin
    
    -- normalize
    input := replace(input, CHAR_CR||CHAR_NL, CHAR_NL);
    input := replace(input, CHAR_CR, CHAR_NL);
    input := replace(input, CHAR_FF, CHAR_NL);
    
    tokenize(input);
    
    decls := declarationList_t();
    
    -- 5.4.5. Consume a list of declarations
    loop
      
      currentToken := consumeNextToken;
      exit when currentToken.t = T_EOF;
      
      if currentToken.t in (T_WHITESPACE,T_SEMICOLON) then
        null;
      elsif currentToken.t = T_IDENT then
      
        decl.name := currentToken.v;
        decl.v := componentList_t();
        decl.argListMap.delete;
        
        while nextToken().t = T_WHITESPACE loop
          consumeNextToken;
        end loop;
        if nextToken().t != T_COLON then
          unexpectedToken(nextToken);
        end if;
        consumeNextToken; -- consume ':'
        while nextToken().t = T_WHITESPACE loop
          consumeNextToken;
        end loop;
        
        inList := false;
        
        -- components
        while nextToken().t not in (T_SEMICOLON,T_EOF) loop
          
          currentToken := consumeNextToken;
          
          if currentToken.t in (T_STRING,T_IDENT,T_DIMENSION,T_PERCENTAGE,T_NUMBER,T_HASH,T_FUNCTION) then
            
            -- read function value
            if currentToken.t = T_FUNCTION then
              consumeFunction(currentToken);
            end if;
          
            if not inList then
              comp.token := currentToken;
              comp.tokenList := null;
              comp.isList := false;
              decl.v.extend;
              decl.v(decl.v.last) := comp;
            else
              comp := decl.v(decl.v.last);
              if comp.tokenList is null then
                comp.tokenList := tokenList_t();
                appendToken(comp.tokenList, comp.token);
                comp.token := null;
                comp.isList := true;
              end if;
              appendToken(comp.tokenList, currentToken);
              decl.v(decl.v.last) := comp;
              inList := false;       
            end if;
            
          elsif currentToken.t = T_WHITESPACE then
            null;
            
          elsif currentToken.t = T_COMMA then
            if inList then
              unexpectedToken(currentToken);
            else
              inList := true;
            end if;
          
          else
            unexpectedToken(currentToken);         
          end if;
        
        end loop;
        
        -- trim whitespaces
        while decl.v.count > 0 and decl.v(decl.v.last).token.t = T_WHITESPACE loop
          decl.v.trim;
        end loop;
        
        if decl.v.count = 0 then
          error('Empty value for property ''%s''', decl.name);
        end if;
        
        if inList then
          error('Unexpected end of value');
        end if;
        
        decls.extend;
        decls(decls.last) := decl;
        
      else
        unexpectedToken(currentToken);
      end if;
      
    end loop;
    
    return decls;
    
  end;
  
  function isValidCssBorderStyle (p_value in varchar2) return boolean is
  begin
    return p_value in ('none','solid','dashed','dotted','double','hairline','dot-dash','dot-dot-dash','dot-dash-slanted');
  end;

  function isValidCssBorderWidth (p_value in varchar2) return boolean is
  begin
    return p_value in ('thin','medium','thick');
  end;
  
  function isValidMsoPatternType (p_value in varchar2) return boolean is
  begin
    return p_value in ('none','gray-50','gray-75','gray-25','horz-stripe','vert-stripe','reverse-dark-down'
                      ,'diag-stripe','diag-cross','thick-diag-cross','thin-horz-stripe','thin-vert-stripe'
                      ,'thin-reverse-diag-stripe','thin-diag-stripe','thin-horz-cross','thin-diag-cross'
                      ,'gray-125','gray-0625');
  end;
  
  function "rgb" (args in tokenList_t) return varchar2 is
    token      token_t;
    argType    pls_integer;
    sepType    pls_integer;
    color      rgbColor_t;
    idx        pls_integer := 0;
    
    procedure nextToken is
    begin
      if idx = args.count then
        token.t := T_EOF;
      else
        idx := idx + 1;
        token := args(idx);
      end if;
    end;
    
    procedure skipWhitespace is
    begin
      if token.t = T_WHITESPACE then
        nextToken;
      end if;      
    end;

    procedure expect (t in pls_integer, v in varchar2 default null) is
    begin
      if token.t = t and (v is null or token.v = v) then
        nextToken;
      else
        unexpectedToken(token);
      end if;      
    end;
    
    function readColorComponent (checkType in boolean default true) return number is
      componentValue  number;
    begin
      if checkType and token.t != argType then
        unexpectedToken(token);
      end if;
      case argType
      when T_NUMBER then
        if token.nv between 0 and 255 then
          componentValue := token.nv;
        else
          error('Invalid RGB component value: %s', token.v);
        end if;        
      when T_PERCENTAGE then
        if token.nv between 0 and 100 then
          componentValue := token.nv * 2.55;
        else
          error('Invalid RGB component value: %s', token.v);
        end if;        
      end case;
      nextToken;
      return componentValue;
    end;
    
  begin
    
    if args is empty then
      error('No arguments found for rgb() function');
    end if;
    
    -- argument type detection
    argType := args(1).t;
    if argType not in (T_NUMBER, T_PERCENTAGE) then
      unexpectedToken(args(1));
    end if;
    
    -- separator type detection
    if args.count > 1 then   
      if args(2).t = T_WHITESPACE then
        if args(3).t = T_COMMA then
          sepType := T_COMMA;
        elsif args(3).t = argType then
          sepType := T_WHITESPACE;
        else
          unexpectedToken(args(3));
        end if;
      elsif args(2).t = T_COMMA then
        sepType := T_COMMA;
      else
        unexpectedToken(args(2));
      end if;
    end if;
    
    -- syntax #1: rgb(R, G, B[, A])
    if sepType = T_COMMA then
    
      nextToken;
      --arg1
      color.r := readColorComponent;
      skipWhitespace;
      expect(sepType);
      skipWhitespace;
      
      --arg2
      color.g := readColorComponent;
      skipWhitespace;
      expect(sepType);
      skipWhitespace;

      --arg3
      color.b := readColorComponent;

      if token.t != T_RIGHT then
        --arg4
        skipWhitespace;
        expect(sepType);
        skipWhitespace;
        if token.t in (T_NUMBER,T_PERCENTAGE) then
          color.a := readColorComponent(false);
        else
          unexpectedToken(token);
        end if;
        expect(T_RIGHT);
      end if;
    
    -- syntax #2: rgb(R G B[ / A])
    elsif sepType = T_WHITESPACE then

      nextToken;
      --arg1
      color.r := readColorComponent;     
      expect(sepType);
      
      --arg2
      color.g := readColorComponent;
      expect(sepType);

      --arg3
      color.b := readColorComponent;    

      if token.t != T_RIGHT then
        --arg4
        skipWhitespace;
        expect(T_DELIM, '/');
        skipWhitespace;
        if token.t in (T_NUMBER,T_PERCENTAGE) then
          color.a := readColorComponent(false);
        else
          unexpectedToken(token);
        end if;
        expect(T_RIGHT);
      end if;
    
    end if;
    
    return makeRgbColor(color.r, color.g, color.b, color.a);
    
  end;
  
  function parseColor (token in token_t, validate in boolean default false) return varchar2 is
    colorCode  varchar2(256);
  begin

    case token.t
    when T_IDENT then
            
      if isValidColorName(token.v) then
        colorCode := '#' || getColorCode(token.v);
      elsif validate then
        error(CSS_INVALID_VALUE, token.v);
      end if;
              
    when T_HASH then
          
      if isValidColorCode(token.v) then
        colorCode := '#' || token.v;
      elsif validate then
        error(CSS_BAD_RGB_COLOR, token.v);
      end if;
    
    when T_FUNCTION then
      
      if token.v in ('rgb','rgba') then
        colorCode := '#' || "rgb"(functionArgListCache(token.argListId));
      elsif validate then
        error('Unsupported or invalid function: ''%s''', token.v);
      end if;
         
    else
      
      if validate then
        unexpectedToken(token);
      end if;
      
    end case;
    
    return colorCode;
    
  end;

  function "linear-gradient" (args in tokenList_t) return linearGradient_t is
    token             token_t;
    gradient          linearGradient_t;
    idx               pls_integer := 0;
    hasHorizontalDir  boolean := false;
    hasVerticalDir    boolean := false;
    horizontalDir     varchar2(256);
    verticalDir       varchar2(256);
    colorStop         colorStop_t;
    
    function assertPercentRange (v in number) return number is
    begin
      if v not between 0 and 100 then
        error('Unsupported percentage value: %d', v);
      end if;
      return v;
    end;
    
    procedure nextToken is
    begin
      if idx = args.count then
        token.t := T_EOF;
      else
        idx := idx + 1;
        token := args(idx);
      end if;
    end;
    
    procedure skipWhitespace is
    begin
      if token.t = T_WHITESPACE then
        nextToken;
      end if;      
    end;

    function match (t in pls_integer, v in varchar2 default null) return boolean is
    begin
      return ( token.t = t and (v is null or token.v = v) );      
    end;

    procedure expect (t in pls_integer, v in varchar2 default null) is
    begin
      if match(t,v) then
        nextToken;
      else
        unexpectedToken(token);
      end if;      
    end;
    
    procedure readColorStop is
    begin
      colorStop.color := parseColor(token);
      if colorStop.color is null then
        unexpectedToken(token);
      end if;
      nextToken;
      skipWhitespace;
      if match(T_PERCENTAGE) then
        colorStop.pct1 := assertPercentRange(token.nv);
        nextToken;
        skipWhitespace;
        if match(T_PERCENTAGE) then
          colorStop.pct2 := assertPercentRange(token.nv);
          nextToken;
          skipWhitespace;
        end if;
      end if;
      gradient.colorStopList.extend;
      gradient.colorStopList(gradient.colorStopList.last) := colorStop;
    end;
    
  begin
    
    if args is empty then
      error('No arguments found for linear-gradient() function');
    end if;
    
    gradient.colorStopList := colorStopList_t();
    
    nextToken;
    
    if match(T_DIMENSION) then
      case token.u
      when 'deg' then
        gradient.angle := token.nv;
      when 'turn' then 
        gradient.angle := 360 * token.nv;
      else
        unexpectedToken(token);
      end case;
      nextToken;
      skipWhitespace;
      expect(T_COMMA);
      
    elsif match(T_IDENT, 'to') then
    
      nextToken;
      skipWhitespace;
      if match(T_IDENT, 'left') or match(T_IDENT, 'right') then
        horizontalDir := token.v;
        hasHorizontalDir := true;
      elsif match(T_IDENT, 'top') or match(T_IDENT, 'bottom') then
        verticalDir := token.v;
        hasVerticalDir := true;
      else
        unexpectedToken(token);
      end if;
      nextToken;
      skipWhitespace;
      -- try to match a 2nd gradient direction
      if not hasHorizontalDir and ( match(T_IDENT, 'left') or match(T_IDENT, 'right') ) then
        horizontalDir := token.v;
        hasHorizontalDir := true;
      elsif not hasVerticalDir and ( match(T_IDENT, 'top') or match(T_IDENT, 'bottom') ) then
        verticalDir := token.v;
        hasVerticalDir := true;
      elsif match(T_COMMA) then
        nextToken;
      else
        unexpectedToken(token);
      end if;
      
      gradient.angle := case when horizontalDir = 'left' and verticalDir = 'top' then 315
                             when horizontalDir = 'left' and verticalDir = 'bottom' then 225
                             when horizontalDir = 'right' and verticalDir = 'bottom' then 135
                             when horizontalDir = 'right' and verticalDir = 'top' then 45
                             when horizontalDir = 'left' then 270
                             when verticalDir = 'bottom' then 180
                             when horizontalDir = 'right' then 90
                             when verticalDir = 'top' then 0
                        end;
      
      if hasHorizontalDir and hasVerticalDir then
        nextToken;
        skipWhitespace;
        expect(T_COMMA);
      end if;
       
    end if;
    
    skipWhitespace;
    
    --read first color-stop
    colorStop := null;
    readColorStop;
    
    loop
      
      expect(T_COMMA);
      skipWhitespace;
    
      colorStop := null;  
      --read color hint
      if match(T_PERCENTAGE) then
        colorStop.colorHint := assertPercentRange(token.nv);
        nextToken;
        skipWhitespace;
        expect(T_COMMA);
      end if;
      --read next color-stop
      skipWhitespace;
      readColorStop;
      exit when match(T_RIGHT);
    end loop;
    
    for i in 1 .. gradient.colorStopList.count loop
      debug(
        utl_lms.format_message(
          'hint=%s color=%s, pct1=%s, pct2=%s'
        , to_char(gradient.colorStopList(i).colorHint)
        , gradient.colorStopList(i).color
        , to_char(gradient.colorStopList(i).pct1)
        , to_char(gradient.colorStopList(i).pct2)
        )
      );
    end loop;
    
    return gradient;
    
  end;
  
  procedure parseCssBorderSide (decl in declaration_t, cssBorderSide in out nocopy cssBorderSide_t) 
  is
    token      token_t;
    hasStyle   boolean := false;
    hasWidth   boolean := false;
    hasColor   boolean := false;
    colorCode  varchar2(256);
  begin
    
    assertValueCount(decl.name, decl.v.count, 3);

    for i in 1 .. decl.v.count loop
      
      assertNotIsList(decl.v(i)); 
      token := decl.v(i).token;
      
      -- is it a valid color token?
      colorCode := parseColor(token);
      if colorCode is not null then
        
        if hasColor then
          error('Duplicate color value');
        end if;
        cssBorderSide.color := colorCode;
        hasColor := true;
      
      elsif token.t = T_IDENT then
      
        if not hasStyle and isValidCssBorderStyle(token.v) then
          cssBorderSide.style := token.v;
          hasStyle := true;
        elsif not hasWidth and isValidCssBorderWidth(token.v) then
          cssBorderSide.width := token.v;
          hasWidth := true;
        else
          error(CSS_INVALID_VALUE, token.v);
        end if;
        
      else
        
        unexpectedToken(token);
      
      end if;
        
    end loop;
    
  end;

  procedure parseCssBorderStyle (
    decl in declaration_t
  , cssBorder in out nocopy cssBorder_t
  , borderSideName in varchar2 default null
  ) 
  is
    maxValueCount  pls_integer := case when borderSideName is not null then 1 else 4 end;
    token          token_t;
  begin
    
    assertValueCount(decl.name, decl.v.count, maxValueCount);

    for i in 1 .. decl.v.count loop
            
      assertNotIsList(decl.v(i));
          
      token := decl.v(i).token;
          
      if token.t = T_IDENT then
            
        if isValidCssBorderStyle(token.v) then
          
          case decl.v.count
          when 1 then
            case borderSideName
            when 'top' then
              cssBorder.top.style := token.v;
            when 'right' then
              cssBorder.right.style := token.v;
            when 'bottom' then
              cssBorder.bottom.style := token.v;
            when 'left' then
              cssBorder.left.style := token.v;
            else
              -- set all four sides at once
              cssBorder.top.style := token.v;
              cssBorder.right.style := token.v;
              cssBorder.bottom.style := token.v;
              cssBorder.left.style := token.v;
            end case;
          when 2 then
            -- set top and bottom, left and right
            if i = 1 then
              cssBorder.top.style := token.v;
              cssBorder.bottom.style := token.v;
            else
              cssBorder.left.style := token.v;
              cssBorder.right.style := token.v;
            end if;
          when 3 then
            -- set top, left and right, bottom
            if i = 1 then
              cssBorder.top.style := token.v;      
            elsif i = 2 then
              cssBorder.left.style := token.v;
              cssBorder.right.style := token.v;
            else
              cssBorder.bottom.style := token.v;
            end if;
          else
            -- set top, right, bottom, left
            if i = 1 then
              cssBorder.top.style := token.v;      
            elsif i = 2 then
              cssBorder.right.style := token.v;
            elsif i = 3 then
              cssBorder.bottom.style := token.v;
            else
              cssBorder.left.style := token.v;
            end if;
          end case;
          
        else
          error(CSS_INVALID_VALUE, token.v);
        end if;
          
      else
        unexpectedToken(token);
      end if;
        
    end loop;
    
  end;

  procedure parseCssBorderWidth (
    decl in declaration_t
  , cssBorder in out nocopy cssBorder_t
  , borderSideName in varchar2 default null
  ) 
  is
    maxValueCount  pls_integer := case when borderSideName is not null then 1 else 4 end;
    token          token_t;
  begin
    
    assertValueCount(decl.name, decl.v.count, maxValueCount);

    for i in 1 .. decl.v.count loop
            
      assertNotIsList(decl.v(i));
          
      token := decl.v(i).token;
          
      if token.t = T_IDENT then
            
        if isValidCssBorderWidth(token.v) then
          
          case decl.v.count
          when 1 then
            case borderSideName
            when 'top' then
              cssBorder.top.width := token.v;
            when 'right' then
              cssBorder.right.width := token.v;
            when 'bottom' then
              cssBorder.bottom.width := token.v;
            when 'left' then
              cssBorder.left.width := token.v;
            else
              -- set all four sides at once
              cssBorder.top.width := token.v;
              cssBorder.right.width := token.v;
              cssBorder.bottom.width := token.v;
              cssBorder.left.width := token.v;
            end case;
          when 2 then
            -- set top and bottom, left and right
            if i = 1 then
              cssBorder.top.width := token.v;
              cssBorder.bottom.width := token.v;
            else
              cssBorder.left.width := token.v;
              cssBorder.right.width := token.v;
            end if;
          when 3 then
            -- set top, left and right, bottom
            if i = 1 then
              cssBorder.top.width := token.v;      
            elsif i = 2 then
              cssBorder.left.width := token.v;
              cssBorder.right.width := token.v;
            else
              cssBorder.bottom.width := token.v;
            end if;
          else
            -- set top, right, bottom, left
            if i = 1 then
              cssBorder.top.width := token.v;      
            elsif i = 2 then
              cssBorder.right.width := token.v;
            elsif i = 3 then
              cssBorder.bottom.width := token.v;
            else
              cssBorder.left.width := token.v;
            end if;
          end case;
          
        else
          error(CSS_INVALID_VALUE, token.v);
        end if;
          
      else
        unexpectedToken(token);
      end if;
        
    end loop;
    
  end;

  procedure parseCssBorderColor (
    decl            in declaration_t
  , cssBorder       in out nocopy cssBorder_t
  , borderSideName  in varchar2 default null
  ) 
  is
    maxValueCount  pls_integer := case when borderSideName is not null then 1 else 4 end;
    colorCode      varchar2(256);
  begin
    
    assertValueCount(decl.name, decl.v.count, maxValueCount);

    for i in 1 .. decl.v.count loop
            
      assertNotIsList(decl.v(i));
      colorCode := parseColor(decl.v(i).token, true);
                
      case decl.v.count
      when 1 then
        case borderSideName
        when 'top' then
          cssBorder.top.color := colorCode;
        when 'right' then
          cssBorder.right.color := colorCode;
        when 'bottom' then
          cssBorder.bottom.color := colorCode;
        when 'left' then
          cssBorder.left.color := colorCode;
        else
          -- set all four sides at once
          cssBorder.top.color := colorCode;
          cssBorder.right.color := colorCode;
          cssBorder.bottom.color := colorCode;
          cssBorder.left.color := colorCode;
        end case;
      when 2 then
        -- set top and bottom, left and right
        if i = 1 then
          cssBorder.top.color := colorCode;
          cssBorder.bottom.color := colorCode;
        else
          cssBorder.left.color := colorCode;
          cssBorder.right.color := colorCode;
        end if;
      when 3 then
        -- set top, left and right, bottom
        if i = 1 then
          cssBorder.top.color := colorCode;      
        elsif i = 2 then
          cssBorder.left.color := colorCode;
          cssBorder.right.color := colorCode;
        else
          cssBorder.bottom.color := colorCode;
        end if;
      else
        -- set top, right, bottom, left
        if i = 1 then
          cssBorder.top.color := colorCode;      
        elsif i = 2 then
          cssBorder.right.color := colorCode;
        elsif i = 3 then
          cssBorder.bottom.color := colorCode;
        else
          cssBorder.left.color := colorCode;
        end if;
      end case;
        
    end loop;
    
  end;

  procedure parseCssFont (
    decl in declaration_t
  , font in out nocopy cssFont_t
  )
  is
    valueCount  pls_integer := decl.v.count;
    
    hasStyle   boolean := false;
    hasWeight  boolean := false;
    
    procedure readFamily (comp in component_t) is
    begin
      assertNotIsList(comp);
      if comp.token.t in (T_IDENT,T_STRING) then
        font.family := comp.token.v;
      else
        unexpectedToken(comp.token);
      end if;
    end;
    
    procedure readSize (comp in component_t) is
    begin
      assertNotIsList(comp);
      if comp.token.t = T_DIMENSION then
        -- 'pt' is the only supported unit
        if comp.token.u != 'pt' then
          error('Unsupported font-size unit: ''%s''', comp.token.u);
        end if;
        font.sz := comp.token.v;
      else
        unexpectedToken(comp.token);
      end if;      
    end;

    function readStyle (comp in component_t) return boolean is
    begin
      assertNotIsList(comp);
      if comp.token.t = T_IDENT then
        if comp.token.v in ('italic','normal') then
          font.style := comp.token.v;
        else
          return false;
        end if;
      else
        unexpectedToken(comp.token);
      end if;
      return true;      
    end;

    function readWeight (comp in component_t) return boolean is
    begin
      assertNotIsList(comp);
      if comp.token.t = T_IDENT then
        if comp.token.v in ('bold','normal') then
          font.weight := comp.token.v;
        else
          return false;
        end if;
      else
        unexpectedToken(comp.token);
      end if;
      return true;      
    end;
    
  begin
    
    if decl.name = 'font' then
      assertValueCount(decl.name, valueCount, 4, 2);
    else
      assertValueCount(decl.name, valueCount, 1);
    end if;
    
    case decl.name 
    when 'font' then
      readFamily(decl.v(valueCount));
      readSize(decl.v(valueCount-1));
      
      if valueCount > 2 then
      
        hasStyle := readStyle(decl.v(1));
        if not hasStyle then
          hasWeight := readWeight(decl.v(1));
        end if;
      
        if not ( hasStyle or hasWeight ) then
          error(CSS_INVALID_VALUE, decl.v(1).token.v);
        end if;
        
        if valueCount > 3 then
          
          -- if 1st value is 'normal', it's been assigned to font-style by design but it may also apply to font-weight
          -- so if 2nd value is a valid font-style, consider 1st value to be a font-weight
          if decl.v(1).token.v = 'normal' and readStyle(decl.v(2)) then
            font.weight := 'normal';
            hasWeight := true;
          end if;
        
          if not hasStyle then
            hasStyle := readStyle(decl.v(2));
          elsif not hasWeight then
            hasWeight := readWeight(decl.v(2));
          end if;
        
          if not ( hasStyle and hasWeight ) then
            error(CSS_INVALID_VALUE, decl.v(2).token.v);
          end if;
          
        end if;
        
      end if;
      
    when 'font-family' then
      readFamily(decl.v(1));
      
    when 'font-size' then
      readSize(decl.v(1));
      
    when 'font-style' then
      if not readStyle(decl.v(1)) then
        error(CSS_INVALID_VALUE, decl.v(1).token.v);
      end if;
      
    when 'font-weight' then
      if not readWeight(decl.v(1)) then
        error(CSS_INVALID_VALUE, decl.v(1).token.v);
      end if;
      
    end case;
    
  end;

  procedure parseCssTextDecoration (
    decl in declaration_t
  , textDecoration in out nocopy cssTextDecoration_t
  )
  is
    valueCount  pls_integer := decl.v.count;
    hasLine     boolean := false;
    hasStyle    boolean := false;
    
    function readLine (comp in component_t) return boolean is
    begin
      assertNotIsList(comp);
      if comp.token.t = T_IDENT then
        if comp.token.v in ('none','underline') then
          textDecoration.line := comp.token.v;
        else
          return false;
        end if;
      else
        unexpectedToken(comp.token);
      end if;
      return true;
    end;

    function readStyle (comp in component_t) return boolean is
    begin
      assertNotIsList(comp);
      if comp.token.t = T_IDENT then
        if comp.token.v in ('solid','double','single-accounting','double-accounting') then
          textDecoration.style := comp.token.v;
        else
          return false;
        end if;
      else
        unexpectedToken(comp.token);
      end if;
      return true;      
    end;
    
  begin
    
    if decl.name = 'text-decoration' then
      assertValueCount(decl.name, valueCount, 2);
    else
      assertValueCount(decl.name, valueCount, 1);
    end if;
    
    case decl.name 
    when 'text-decoration' then
      
      hasLine := readLine(decl.v(1));
      if not hasLine then
        hasStyle := readStyle(decl.v(1));
      end if;
    
      if not ( hasLine or hasStyle ) then
        error(CSS_INVALID_VALUE, decl.v(1).token.v);
      end if;
      
      if valueCount > 1 then
        
        if not hasLine then
          hasLine := readLine(decl.v(2));
        elsif not hasStyle then
          hasStyle := readStyle(decl.v(2));
        end if;
      
        if not ( hasLine and hasStyle ) then
          error(CSS_INVALID_VALUE, decl.v(2).token.v);
        end if;
        
      end if;
      
    when 'text-decoration-line' then
      if not readLine(decl.v(1)) then
        error(CSS_INVALID_VALUE, decl.v(1).token.v);
      end if;
      
    when 'text-decoration-style' then
      if not readStyle(decl.v(1)) then
        error(CSS_INVALID_VALUE, decl.v(1).token.v);
      end if;
      
    end case;
    
  end;

  procedure parseCssVerticalAlign (
    decl           in declaration_t
  , verticalAlign  in out nocopy varchar2
  )
  is
    token  token_t;
  begin
    
    assertValueCount(decl.name, decl.v.count, 1);
    assertNotIsList(decl.v(1));
    
    token := decl.v(1).token;
    
    if token.t = T_IDENT then
      
      if token.v in ('top','middle','bottom','justify','distributed') then
        verticalAlign := token.v;
      else
        error(CSS_INVALID_VALUE, token.v);
      end if;
      
    else
      
      unexpectedToken(token);
    
    end if;
    
  end;

  procedure parseCssTextAlign (
    decl       in declaration_t
  , textAlign  in out nocopy varchar2
  )
  is
    token  token_t;
  begin
    
    assertValueCount(decl.name, decl.v.count, 1);
    assertNotIsList(decl.v(1));
    
    token := decl.v(1).token;
    
    if token.t = T_IDENT then
      
      if token.v in ('left','center','right','fill','justify','center-across','distributed') then
        textAlign := token.v;
      else
        error(CSS_INVALID_VALUE, token.v);
      end if;
      
    else
      
      unexpectedToken(token);
    
    end if;
    
  end;

  procedure parseCssWhiteSpace (
    decl        in declaration_t
  , whiteSpace  in out nocopy varchar2
  )
  is
    token  token_t;
  begin
    
    assertValueCount(decl.name, decl.v.count, 1);
    assertNotIsList(decl.v(1));
    
    token := decl.v(1).token;
    
    if token.t = T_IDENT then
      
      if token.v in ('pre','pre-wrap') then
        whiteSpace := token.v;
      else
        error(CSS_INVALID_VALUE, token.v);
      end if;
      
    else
      
      unexpectedToken(token);
    
    end if;
    
  end;

  procedure parseCssColor (
    decl       in declaration_t
  , colorCode  in out nocopy varchar2
  )
  is
  begin
    assertValueCount(decl.name, decl.v.count, 1);
    assertNotIsList(decl.v(1));
    colorCode := parseColor(decl.v(1).token, true);
  end;

  procedure parseCssMsoPattern (
    decl        in declaration_t
  , msoPattern  in out nocopy cssMsoPattern_t
  )
  is
    hasPatternType  boolean := false;
    hasColor        boolean := false;
    colorCode       varchar2(256);
    token           token_t;
  begin
    
    assertValueCount(decl.name, decl.v.count, 2);
    
    for i in 1 .. decl.v.count loop
      
      assertNotIsList(decl.v(i));     
      token := decl.v(i).token;
      
      -- is it a valid color token?
      colorCode := parseColor(token);
      if colorCode is not null then
        
        if hasColor then
          error('Duplicate color value');
        end if;
        msoPattern.color := colorCode;
        hasColor := true;
      
      elsif token.t = T_IDENT then
            
        if not hasPatternType and isValidMsoPatternType(token.v) then
          msoPattern.patternType := token.v;
          hasPatternType := true;
        else
          error(CSS_INVALID_VALUE, token.v);
        end if;
          
      else
        
        unexpectedToken(token);
        
      end if;
        
    end loop;
    
    if not hasPatternType then
      msoPattern.patternType := null;
    end if;
    
  end;

  procedure parseCssMsoNumberFormat (
    decl             in declaration_t
  , msoNumberFormat  in out nocopy varchar2
  )
  is
    token  token_t;
  begin
    
    assertValueCount(decl.name, decl.v.count, 1);
    assertNotIsList(decl.v(1));
    
    token := decl.v(1).token;
    
    if token.t = T_STRING then     
      msoNumberFormat := token.v;
    else
      unexpectedToken(token);
    end if;
    
  end;

  procedure parseCssGradient (
    decl      in declaration_t
  , gradient  in out nocopy linearGradient_t
  )
  is
    token  token_t;
  begin
    assertValueCount(decl.name, decl.v.count, 1);
    assertNotIsList(decl.v(1));
    token := decl.v(1).token;
    
    if token.t = T_FUNCTION then
      
      if token.v = 'linear-gradient' then
        gradient := "linear-gradient"(functionArgListCache(decl.v(1).token.argListId));
      else
        error('Unsupported or invalid function: ''%s''', token.v);
      end if;
      
    else
      unexpectedToken(token);
    end if;
    
  end;
  
  function parseCss (
    decls in declarationList_t
  )
  return cssStyle_t
  is
    decl           declaration_t;
    css            cssStyle_t;    
    cssBorderSide  cssBorderSide_t;
  begin

    for i in 1 .. decls.count loop
    
      decl := decls(i);
      functionArgListCache := decl.argListMap;
    
      case decl.name
      when 'border' then
        
        parseCssBorderSide(decl, cssBorderSide);
        css.border.top := cssBorderSide;
        css.border.right := cssBorderSide;
        css.border.bottom := cssBorderSide;
        css.border.left := cssBorderSide;
        
      when 'border-left' then
         
        parseCssBorderSide(decl, css.border.left);
        
      when 'border-right' then
         
        parseCssBorderSide(decl, css.border.right);
        
      when 'border-top' then
         
        parseCssBorderSide(decl, css.border.top);
        
      when 'border-bottom' then
         
        parseCssBorderSide(decl, css.border.bottom);
        
      when 'border-style' then
        
        parseCssBorderStyle(decl, css.border);
        
      when 'border-width' then
        
        parseCssBorderWidth(decl, css.border);
        
      when 'border-color' then
        
        parseCssBorderColor(decl, css.border);
        
      when 'border-top-style' then
        
        parseCssBorderStyle(decl, css.border, 'top');

      when 'border-right-style' then
        
        parseCssBorderStyle(decl, css.border, 'right');
        
      when 'border-bottom-style' then
        
        parseCssBorderStyle(decl, css.border, 'bottom');
        
      when 'border-left-style' then
        
        parseCssBorderStyle(decl, css.border, 'left');

      when 'border-top-width' then
        
        parseCssBorderWidth(decl, css.border, 'top');

      when 'border-right-width' then
        
        parseCssBorderWidth(decl, css.border, 'right');
        
      when 'border-bottom-width' then
        
        parseCssBorderWidth(decl, css.border, 'bottom');
        
      when 'border-left-width' then
        
        parseCssBorderWidth(decl, css.border, 'left');

      when 'border-top-color' then
        
        parseCssBorderColor(decl, css.border, 'top');

      when 'border-right-color' then
        
        parseCssBorderColor(decl, css.border, 'right');
        
      when 'border-bottom-color' then
        
        parseCssBorderColor(decl, css.border, 'bottom');
        
      when 'border-left-color' then
        
        parseCssBorderColor(decl, css.border, 'left');
        
      when 'font' then
        
        parseCssFont(decl, css.font);
        
      when 'font-family' then
        
        parseCssFont(decl, css.font);
      
      when 'font-size' then
        
        parseCssFont(decl, css.font);
        
      when 'font-style' then
        
        parseCssFont(decl, css.font);
        
      when 'font-weight' then
        
        parseCssFont(decl, css.font);
        
      when 'text-decoration' then
        
        parseCssTextDecoration(decl, css.textDecoration);

      when 'text-decoration-line' then
        
        parseCssTextDecoration(decl, css.textDecoration);
        
      when 'text-decoration-style' then
        
        parseCssTextDecoration(decl, css.textDecoration);
        
      when 'vertical-align' then
        
        parseCssVerticalAlign(decl, css.verticalAlign);
        
      when 'text-align' then
        
        parseCssTextAlign(decl, css.textAlign);
        
      when 'white-space' then
        
        parseCssWhiteSpace(decl, css.whiteSpace);
      
      when 'color' then
        
        parseCssColor(decl, css.color);
        
      when 'background' then
        
        begin
          parseCssColor(decl, css.backgroundColor);
        exception
          when others then
            parseCssGradient(decl, css.backgroundImage);
        end;

      when 'background-color' then
        
        parseCssColor(decl, css.backgroundColor);
        
      when 'background-image' then
        
        parseCssGradient(decl, css.backgroundImage);
        
      when 'mso-number-format' then
        
        parseCssMsoNumberFormat(decl, css.msoNumberFormat);
      
      when 'mso-pattern' then
        
        parseCssMsoPattern(decl, css.msoPattern);
        
      else
        
        error('Unsupported CSS property: ''%s''', decl.name);
        
      end case;
    
    end loop;
    
    return css;
    
  end;
      
  function convertCssBorderSide (cssBorderSide in cssBorderSide_t) return CT_BorderPr is
    borderPr  CT_BorderPr;
  begin
  
    case cssBorderSide.style
    when 'none' then
      borderPr.style := 'none';
    when 'solid' then
      borderPr.style := case cssBorderSide.width
                        when 'thin' then 'thin'
                        when 'medium' then 'medium'
                        when 'thick' then 'thick'
                        end ;
    when 'dashed' then
      borderPr.style := case cssBorderSide.width
                        when 'thin' then 'dashed'
                        when 'medium' then 'mediumDashed'
                        else 'mediumDashed'
                        end ;
    when 'dotted' then
      borderPr.style := 'dotted';
    when 'double' then
      borderPr.style := 'double';
    when 'hairline' then
      borderPr.style := 'hair';
    when 'dot-dash' then
      borderPr.style := case cssBorderSide.width
                        when 'thin' then 'dashDot'
                        when 'medium' then 'mediumDashDot'
                        else 'mediumDashDot'
                        end ;
    when 'dot-dot-dash' then
      borderPr.style := case cssBorderSide.width
                        when 'thin' then 'dashDotDot'
                        when 'medium' then 'mediumDashDotDot'
                        else 'mediumDashDotDot'
                        end ;
    when 'dot-dash-slanted' then
      borderPr.style := 'slantDashDot';
    end case;
    
    borderPr.color := validateColor(cssBorderSide.color);
    
    return borderPr;
    
  end;

  function convertCssBorder (cssBorder in cssBorder_t) return CT_Border is
  begin
    return makeBorder( convertCssBorderSide(cssBorder.left)
                     , convertCssBorderSide(cssBorder.right)
                     , convertCssBorderSide(cssBorder.top)
                     , convertCssBorderSide(cssBorder.bottom) );
  end;

  

  function convertCssFont (css in cssStyle_t) return CT_Font is
    font  CT_Font;
  begin
    
    font.name := css.font.family;
    font.sz := regexp_substr(css.font.sz, '^\d+'); -- digits only
    font.i := ( css.font.style = 'italic' );
    font.b := ( css.font.weight = 'bold' );
    
    if css.textDecoration.line = 'underline' then
      font.u := case css.textDecoration.style
                when 'solid' then 'single'
                when 'double' then 'double'
                when 'single-accounting' then 'singleAccounting'
                when 'double-accounting' then 'doubleAccounting'
                end;
    else
      font.u := 'none';
    end if;
    
    font.color := validateColor(css.color);
    
    setFontContent(font);
    
    return font;
  
  end;

  function convertCssAlign (css in cssStyle_t) return CT_CellAlignment is
    alignment  CT_CellAlignment;
  begin
  
    alignment.horizontal := case css.textAlign
                            when 'left' then 'left'
                            when 'center' then 'center'
                            when 'right' then 'right'
                            when 'fill' then 'fill'
                            when 'justify' then 'justify'
                            when 'center-across' then 'centerContinuous'
                            when 'distributed' then 'distributed'
                            end;
                            
    alignment.vertical := case css.verticalAlign
                          when 'top' then 'top'
                          when 'middle' then 'center'
                          when 'bottom' then 'bottom'
                          when 'justify' then 'justify'
                          when 'distributed' then 'distributed'
                          end;
    
    alignment.wrapText := ( css.whiteSpace = 'pre-wrap' );
    
    setAlignmentContent(alignment);
    
    return alignment;
  
  end;

  function convertCssGradient (gradient in linearGradient_t) return CT_GradientFill is
    gradientFill  CT_GradientFill;
    tmp           CT_GradientStopList := CT_GradientStopList();
    stops         colorStopList_t := gradient.colorStopList;
    sofar         number;
    idx           pls_integer;
    n             pls_integer;
    len           number;
    procedure putStop (pos in number, color in varchar2) is
    begin
      tmp.extend;
      tmp(tmp.last) := makeGradientStop(pos/100, color);
    end;
  begin
    -- default orientation in CSS linear-gradient is top-bottom, which corresponds to 180deg in CSS convention
    -- Excel's default is left-right, which corresponds to 0deg in Excel convention
    -- so Excel degree = CSS angle - 90
    gradientFill.degree := mod(nvl(gradient.angle,180) - 90, 360);
    gradientFill.stops := CT_GradientStopList();
    
    -- https://drafts.csswg.org/css-images-4/#gradient-colors
    -- 3.5.3. Color Stop "Fixup"
    -- TODO: move to "linear-gradient()"?
    
    /* 1a. If the first color stop does not have a position, set its position to 0%. */
    if stops(1).pct1 is null then
      stops(1).pct1 := 0;
    end if;
    /* 1b. If the last color stop does not have a position, set its position to 100%. */
    if stops(stops.last).pct1 is null then
      stops(stops.last).pct1 := 100;
    end if;
    
    /* 2. If a color stop or transition hint has a position that is less than the specified position 
          of any color stop or transition hint before it in the list, set its position to be equal to 
          the largest specified position of any color stop or transition hint before it. */
    sofar := 0;
    for i in 2 .. stops.count loop
      
      if stops(i).colorHint is not null then
        if stops(i).colorHint < sofar then
          stops(i).colorHint := sofar;
        else
          sofar := stops(i).colorHint;
        end if; 
      end if;
      
      if stops(i).pct1 is not null then
        if stops(i).pct1 < sofar then
          stops(i).pct1 := sofar;
        else
          sofar := stops(i).pct1;
        end if; 
      end if;

      if stops(i).pct2 is not null then
        if stops(i).pct2 < sofar then
          stops(i).pct2 := sofar;
        else
          sofar := stops(i).pct2;
        end if; 
      end if;      
    
    end loop;
    
    /* 3. If any color stop still does not have a position, then, for each run of adjacent color stops 
          without positions, set their positions so that they are evenly spaced between the preceding 
          and following color stops with positions. */
              
    for i in 1 .. stops.count - 1 loop
      
      if stops(i).pct1 is null then
        
        -- get next stop with position
        idx := i + 1;
        while stops(idx).pct1 is null loop
          idx := idx + 1;
        end loop;
        
        -- set empty positions
        n := idx - i + 1;
        len := (stops(idx).pct1 - sofar)/n;
        for j in 0 .. n-2 loop
          stops(i+j).pct1 := sofar + len * (j+1);
        end loop;
        
      else
        
        sofar := stops(i).pct1;
        if stops(i).pct2 is not null then
          sofar := stops(i).pct2;
        end if;
      
      end if;
    
    end loop;
    
    -- first stop
    putStop(0, stops(1).color);
    if stops(1).pct2 > stops(1).pct1 then
      putStop(stops(1).pct2, stops(1).color);
    elsif stops(1).pct1 > 0 then
      putStop(stops(1).pct1, stops(1).color);
    end if;
    
    for i in 2 .. stops.count - 1 loop
      
      -- transition hint
      if stops(i).colorHint is not null then
        putStop(stops(i).colorHint, mixRgbColor(stops(i-1).color, stops(i).color));
      end if;
      
      putStop(stops(i).pct1, stops(i).color);
      if stops(i).pct2 != stops(i).pct1 then
        putStop(stops(i).pct2, stops(i).color);
      end if;
    
    end loop;

    -- last transition hint
    if stops(stops.last).colorHint is not null then
      putStop(stops(stops.last).colorHint, '#'||mixRgbColor(substr(stops(stops.last-1).color,2), substr(stops(stops.last).color,2)));
    end if;
    
    -- last stop
    if stops(stops.last).pct1 < 100 then
      putStop(stops(stops.last).pct1, stops(stops.last).color);
    end if;    
    putStop(100, stops(stops.last).color);
    
    -- cleanup
    -- discard stops before the last leading 0 and after the first 1 positions
    idx := 1;
    while tmp(idx).position = 0 loop
      idx := idx + 1;
    end loop;
    
    for i in idx-1 .. tmp.last loop
      gradientFill.stops.extend;
      gradientFill.stops(gradientFill.stops.last) := tmp(i);
      exit when tmp(i).position = 1;
    end loop;
    
    -- Excel doesn't handle identical stop positions the same way as CSS
    -- so adding epsilon increments to distinguish them internally, yet keeping them visually at the same position
    sofar := 0;
    n := 0;
    for i in 2 .. gradientFill.stops.count - 1 loop
      if gradientFill.stops(i).position = sofar then
        n := n + 1;
        gradientFill.stops(i).position := sofar + n * 1e-15;
      else
        n := 0;
        sofar := gradientFill.stops(i).position;
      end if;
    end loop;
  
    return gradientFill;
  end;

  function convertCssBackground (css in cssStyle_t) return CT_Fill is
    fill  CT_Fill;
  begin
    
    if css.backgroundImage.colorStopList is null then
    
      if css.msoPattern.patternType = 'none' then
      
        --fill.patternFill.patternType := 'solid';
        fill.patternFill.fgColor := validateColor(css.backgroundColor);
        fill.patternFill.bgColor := validateColor(css.msoPattern.color);
        
        if fill.patternFill.fgColor is not null or fill.patternFill.bgColor is not null then
          fill.patternFill.patternType := 'solid';
        else
          fill.patternFill.patternType := 'none';
        end if;
        
      elsif css.msoPattern.patternType is not null then
      
        fill.patternFill.fgColor := validateColor(css.msoPattern.color);
        fill.patternFill.bgColor := validateColor(css.backgroundColor);
      
        case css.msoPattern.patternType
        when 'gray-50' then
          fill.patternFill.patternType := 'mediumGray';
        when 'gray-75' then
          fill.patternFill.patternType := 'darkGray';
        when 'gray-25' then
          fill.patternFill.patternType := 'lightGray';
        when 'horz-stripe' then
          fill.patternFill.patternType := 'darkHorizontal';
        when 'vert-stripe' then
          fill.patternFill.patternType := 'darkVertical';
        when 'reverse-dark-down' then
          fill.patternFill.patternType := 'darkDown';
        when 'diag-stripe' then
          fill.patternFill.patternType := 'darkUp';
        when 'diag-cross' then
          fill.patternFill.patternType := 'darkGrid';
        when 'thick-diag-cross' then
          fill.patternFill.patternType := 'darkTrellis';
        when 'thin-horz-stripe' then
          fill.patternFill.patternType := 'lightHorizontal';
        when 'thin-vert-stripe' then
          fill.patternFill.patternType := 'lightVertical';
        when 'thin-reverse-diag-stripe' then
          fill.patternFill.patternType := 'lightDown';
        when 'thin-diag-stripe' then
          fill.patternFill.patternType := 'lightUp';
        when 'thin-horz-cross' then
          fill.patternFill.patternType := 'lightGrid';
        when 'thin-diag-cross' then
          fill.patternFill.patternType := 'lightTrellis';
        when 'gray-125' then
          fill.patternFill.patternType := 'gray125';
        when 'gray-0625' then
          fill.patternFill.patternType := 'gray0625';
        end case;
    
      else
        fill.patternFill.patternType := 'none';
        fill.patternFill.fgColor := validateColor(css.msoPattern.color);
        fill.patternFill.bgColor := validateColor(css.backgroundColor);
      end if;
      
      fill.fillType := FT_PATTERN;
      setFillContent(fill);
    
    else
      
      fill.gradientFill := convertCssGradient(css.backgroundImage);
      fill.fillType := FT_GRADIENT;
      setFillContent(fill);
    
    end if;
  
    return fill;
    
  end;

  function convertCss (css in cssStyle_t) return CT_Style is
    style  CT_Style;
  begin
    style.numberFormat := css.msoNumberFormat;
    style.font := convertCssFont(css);
    style.fill := convertCssBackground(css);
    style.border := convertCssBorder(css.border);
    style.alignment := convertCssAlign(css);
    return style;
  end;
  
  function getStyleFromCss (cssString in varchar2) return CT_Style is
  begin
    return convertCss(parseCss(parseRawCss(cssString)));
  end;
  
  procedure testCss (cssString in varchar2) is
    style  CT_Style := getStyleFromCss(cssString);
  begin
    dbms_output.put_line(style.numberFormat);
    dbms_output.put_line(style.font.content);
    dbms_output.put_line(style.fill.content);
    dbms_output.put_line(style.border.content);
    dbms_output.put_line(style.alignment.content);
  end;

begin
  
  initialize;

end;
/
