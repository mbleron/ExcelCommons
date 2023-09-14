create or replace type body ExcelVariant is

  member function getNumber (format in varchar2 default null, nullOnError in pls_integer default 0) return number
  is
    num  number;
  begin
    if value is not null then
      case value.getTypeName()
      when 'SYS.NUMBER' then 
        num := value.accessNumber();
      when 'SYS.VARCHAR2' then 
        begin
          if format is not null then
            num := to_number(value.accessVarchar2(), format);
          else
            num := to_number(value.accessVarchar2());
          end if;
        exception
          when others then
            if nvl(nullOnError,0) = 0 then
              raise;
            end if;
        end;
      else
        null;
      end case;
    end if;
    return num;
  end;

  member function getString return varchar2
  is
  begin
    if value is not null then
      return case value.getTypeName()
             when 'SYS.VARCHAR2' then value.accessVarchar2()
             when 'SYS.NUMBER' then to_char(value.accessNumber())
             when 'SYS.DATE' then to_char(value.accessDate())
             when 'SYS.TIMESTAMP' then to_char(value.accessTimestamp())
             when 'SYS.CLOB' then dbms_lob.substr(value.accessClob())
             end;
    else
      return null;
    end if;
  end;

  member function getDate (format in varchar2 default null, nullOnError in pls_integer default 0) return date
  is
    dt  date;
    ts  timestamp;
  begin
    if value is not null then
      case value.getTypeName()
      when 'SYS.DATE' then 
        dt := value.accessDate();
      when 'SYS.TIMESTAMP' then 
        ts := value.accessTimestamp();
        dt := cast(ts + numtodsinterval(round(extract(second from ts)) - extract(second from ts), 'second') as date);
      when 'SYS.VARCHAR2' then
        begin
          dt := to_date(value.accessVarchar2(), nvl(format, sys_context('userenv','nls_date_format')));
        exception
          when others then
            if nvl(nullOnError,0) = 0 then
              raise;
            end if;
        end;
      else
        null;
      end case;
    end if;
    return dt;
  end;

  member function getTimestamp (format in varchar2 default null, nullOnError in pls_integer default 0) return timestamp
  is
    ts  timestamp;
  begin
    if value is not null then
      case value.getTypeName()
      when 'SYS.TIMESTAMP' then
        ts := value.accessTimestamp();
      when 'SYS.DATE' then 
        ts := cast(value.accessDate() as timestamp);
      when 'SYS.VARCHAR2' then
        begin
          if format is not null then
            ts := to_timestamp(value.accessVarchar2(), format);
          else
            ts := to_timestamp(value.accessVarchar2());
          end if;
        exception
          when others then
            if nvl(nullOnError,0) = 0 then
              raise;
            end if;
        end;
      else
        null;
      end case;
    end if;
    return ts;
  end;

  member function getClob return clob
  is
  begin
    if value is not null then
      return case value.getTypeName()
             when 'SYS.CLOB' then value.accessClob()
             when 'SYS.VARCHAR2' then to_clob(value.accessVarchar2())
             end;
    else
      return null;
    end if;
  end;

end;
/
