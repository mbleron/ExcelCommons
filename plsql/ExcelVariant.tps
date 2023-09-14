create or replace type ExcelVariant force as object (
  value anydata
, member function getNumber (format in varchar2 default null, nullOnError in pls_integer default 0) return number
, member function getString return varchar2
, member function getDate (format in varchar2 default null, nullOnError in pls_integer default 0) return date
, member function getTimestamp (format in varchar2 default null, nullOnError in pls_integer default 0) return timestamp
, member function getClob return clob
)
/
