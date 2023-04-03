CREATE OR REPLACE PACKAGE PCK_DAILY_REPORT
AS
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= PROC RUN QUERY (PASSING SQL QUERY )
---DATE :09MAY2019
---PROCEDURE -SPECIFICATION- p_run_sql_query
**************************************/
 PROCEDURE p_run_sql_query(p_sql_query IN VARCHAR2);
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:=START WORKBOOK
---DATE :09MAY2019
---PROCEDURE -SPECIFICATION- p_workbook_start
**************************************/
PROCEDURE p_workbook_start;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:=END WORKBOOK
---DATE :09MAY2019
---PROCEDURE - SPECIFICATION-p_worksheet_end
**************************************/
PROCEDURE p_workbook_end;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= PROC RUN QUERY (PASSING SQL QUERY )
---DATE :09MAY2019
---procedure SPECIFICATION :p_worksheet_start
**************************************/
PROCEDURE p_worksheet_start (p_sheet_name IN VARCHAR2);
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= PROC RUN QUERY (PASSING SQL QUERY )
---DATE :09MAY2019
---procedure BODY : p_worksheet_end
**************************************/
PROCEDURE p_worksheet_end;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:=DATA QUALITY AND STYLE
---DATE :09MAY2019
---PROCEDURE -SPECIFICATION- p_date_style__set
**************************************/
PROCEDURE p_date_style_set;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:=DATA QUALITY AND STYLE
---DATE :09MAY2019
---PROCEDURE -SPECIFICATION- p_date_style__set
**************************************/
PROCEDURE p_call_daily_report;
END PCK_DAILY_REPORT;
/
CREATE OR REPLACE PACKAGE BODY PCK_DAILY_REPORT
AS
/*************************************
---DECLARING VARIABLE (LOCAL + GLOBAL)
**************************************/
 /*Directory Name*/
  v_dir       VARCHAR2(30) := 'DIR_DAILY';
   /*FILE_Name*/
  V_File      Varchar2(30) := 'DAILY_REPORT.xls';
 /*local variable*/
  v_fh        UTL_FILE.FILE_TYPE;
  V_Amount  Integer;
  V_Src_Loc Bfile;
  v_b       BLOB;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= PROC Define the sql to run with varchar
---DATE :09MAY2019
---procedure BODY :p_run_sql_query
**************************************/
PROCEDURE p_run_sql_query(p_sql_query IN VARCHAR2)
IS
    v_v_val     VARCHAR2(4000);
    v_n_val     NUMBER;
    v_d_val     DATE;
    v_ret       NUMBER;
    c           NUMBER;
    d           NUMBER;
    col_cnt     INTEGER;
    f           BOOLEAN;
    rec_tab     DBMS_SQL.DESC_TAB;
    col_num     NUMBER;
  BEGIN
    c := DBMS_SQL.OPEN_CURSOR;
    -- parse the SQL statement
    DBMS_SQL.PARSE(c, p_sql_query, DBMS_SQL.NATIVE);
    -- start execution of the SQL statement
    d := DBMS_SQL.EXECUTE(c);
    -- get a description of the returned columns
    DBMS_SQL.DESCRIBE_COLUMNS(c, col_cnt, rec_tab);
    -- bind variables to columns
    FOR j in 1..col_cnt
    LOOP
      CASE rec_tab(j).col_type
        WHEN 1 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,4000);
        WHEN 2 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_n_val);
        WHEN 12 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_d_val);
      ELSE
        DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,4000);
      END CASE;
    END LOOP;
    -- Output the column headers
    UTL_FILE.PUT_LINE(v_fh,'<ss:Row>');
    FOR j in 1..col_cnt
    LOOP
      UTL_FILE.PUT_LINE(v_fh,'<ss:Cell>');
      UTL_FILE.PUT_LINE(v_fh,'<ss:Data ss:Type="String">'||rec_tab(j).col_name||'</ss:Data>');
      UTL_FILE.PUT_LINE(v_fh,'</ss:Cell>');
    END LOOP;
    UTL_FILE.PUT_LINE(v_fh,'</ss:Row>');
    -- Output the data
    LOOP
      v_ret := DBMS_SQL.FETCH_ROWS(c);
      EXIT WHEN v_ret = 0;
      UTL_FILE.PUT_LINE(v_fh,'<ss:Row>');
      FOR j in 1..col_cnt
      LOOP
        CASE rec_tab(j).col_type
          WHEN 1 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_v_val);
                      UTL_FILE.PUT_LINE(v_fh,'<ss:Cell>');
                      UTL_FILE.PUT_LINE(v_fh,'<ss:Data ss:Type="String">'||v_v_val||'</ss:Data>');
                      UTL_FILE.PUT_LINE(v_fh,'</ss:Cell>');
          WHEN 2 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_n_val);
                      UTL_FILE.PUT_LINE(v_fh,'<ss:Cell>');
                      UTL_FILE.PUT_LINE(v_fh,'<ss:Data ss:Type="Number">'||to_char(v_n_val)||'</ss:Data>');
                      UTL_FILE.PUT_LINE(v_fh,'</ss:Cell>');
          WHEN 12 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_d_val);
                      UTL_FILE.PUT_LINE(v_fh,'<ss:Cell ss:StyleID="OracleDate">');
                      UTL_FILE.PUT_LINE(v_fh,'<ss:Data ss:Type="DateTime">'||to_char(v_d_val,'YYYY-MM-DD"T"HH24:MI:SS')||'</ss:Data>');
                      UTL_FILE.PUT_LINE(v_fh,'</ss:Cell>');
        ELSE
          DBMS_SQL.COLUMN_VALUE(c,j,v_v_val);
          UTL_FILE.PUT_LINE(v_fh,'<ss:Cell>');
          UTL_FILE.PUT_LINE(v_fh,'<ss:Data ss:Type="String">'||v_v_val||'</ss:Data>');
          UTL_FILE.PUT_LINE(v_fh,'</ss:Cell>');
        END CASE;
      END LOOP;
      UTL_FILE.PUT_LINE(v_fh,'</ss:Row>');
    END LOOP;
    DBMS_SQL.CLOSE_CURSOR(c);
 EXCEPTION
     WHEN OTHERS THEN
     DBMS_SQL.CLOSE_CURSOR(c);
 END p_run_sql_query;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= workbook start writing
---DATE :09MAY2019
---procedure BODY :p_workbook_start
**************************************/
PROCEDURE  p_workbook_start
IS
  BEGIN
    UTL_FILE.PUT_LINE(v_fh,'<?xml version="1.0"?>');
    UTL_FILE.PUT_LINE(v_fh,'<ss:Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">');
END p_workbook_start;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= workbook end writing
---DATE :09MAY2019
---procedure BODY :p_workbook_end
**************************************/
PROCEDURE p_workbook_end
IS
  BEGIN
    UTL_FILE.PUT_LINE(v_fh,'</ss:Workbook>');
END p_workbook_end;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= worksheet start writing
---DATE :09MAY2019
---procedure BODY :p_worksheet_start
**************************************/
PROCEDURE p_worksheet_start (p_sheet_name IN VARCHAR2)
  IS
 BEGIN
    UTL_FILE.PUT_LINE(v_fh,'<ss:Worksheet ss:Name="'||p_sheet_name||'">');
    UTL_FILE.PUT_LINE(v_fh,'<ss:Table>');
END p_worksheet_start;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= worksheet END writing
---DATE :09MAY2019
---procedure BODY : p_worksheet_end
**************************************/
PROCEDURE p_worksheet_end IS
  BEGIN
    UTL_FILE.PUT_LINE(v_fh,'</ss:Table>');
    UTL_FILE.PUT_LINE(v_fh,'</ss:Worksheet>');
END p_worksheet_end;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= set data style
---DATE :09MAY2019
---procedure BODY : p_date_style__set
**************************************/
PROCEDURE p_date_style_set 
 IS
  BEGIN
    UTL_FILE.PUT_LINE(v_fh,'<ss:Styles>');
    UTL_FILE.PUT_LINE(v_fh,'<ss:Style ss:ID="OracleDate">');
    UTL_FILE.PUT_LINE(v_fh,'<ss:NumberFormat ss:Format="dd/mm/yyyy\ hh:mm:ss"/>');
    UTL_FILE.PUT_LINE(v_fh,'</ss:Style>');
    UTL_FILE.PUT_LINE(v_fh,'</ss:Styles>');
END p_date_style_set;
/*************************************
---AUTHOR:= Sachin K Singh
---PROPOSE:= Call all the procedure to get the morning report
---DATE :09MAY2019
---procedure BODY : p_call_daily_report
**************************************/
PROCEDURE p_call_daily_report
IS
  BEGIN
    v_fh := UTL_FILE.FOPEN(upper(v_dir),v_file,'w',32767);
    p_workbook_start;
    p_date_style_set;
    p_worksheet_start('TAB1_EMP');
    p_run_sql_query('select * from employees where rownum<10');
    p_worksheet_end;
    p_worksheet_start('TAB2_DEPT');
    p_run_sql_query('select * from departments where rownum<10');
    p_worksheet_end;
    p_workbook_end;
    UTL_FILE.FCLOSE(v_fh);
END p_call_daily_report;
END PCK_DAILY_REPORT;
/
