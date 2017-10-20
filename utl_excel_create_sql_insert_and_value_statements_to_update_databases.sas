Create SQL "insert into and value statements" to insert data into existing Excel(tables,sheets), MySQL or Oracle...;


   WORKING CODE
   ============

     ods tagsets.sql file="d:/sql/class.sql";
     proc print data=df ;

     OUPTUT in d:/sql/class.sql

     Create table class(name varchar(), sex varchar(), age float, height float, weight float);
     Insert into class(name, sex, age, height, weight) Values('Alfred', 'M', 14, 69.0, 112.5);
     Insert into class(name, sex, age, height, weight) Values('Alice', 'F', 13, 56.5, 84.0);

     OUPTUT Can be used to create a SAS dataset or update an existing excel sheet

     UPDATE EXCEL (inplace)

       connect to excel (Path="d:/xls/class.xlsx");
        execute(insert into class values( 20,'Bill','X',15,86.5,132)) by excel;
        execute(insert into class values( 21,'Val', 'I',15,86.5,132)) by excel;

     CREATE INSERT EQUIVALENT OF SAS DATASET (slight edit of output)

     Create table class(name varchar(), sex varchar(), age float, height float, weight float);
     Insert into class(name, sex, age, height, weight)
       Values('Alfred', 'M', 14, 69.0, 112.5)
       Values('Alice', 'F', 13, 56.5, 84.0)
       Values('Barbara', 'F', 13, 65.3, 98.0)
     ;quit;


HAVE workbook d:/xls/class.xlsx with sheet class (you can use [sheet1])

   d:/xls/class.xlsx

      +-----------------------------------------------------------------------------+
      |     A      |     B      |    C       |     D      |    E       |    F       |
      +-----------------------------------------------------------------------------+
   1  |   UID      | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+------------+
   2  |    1       | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+------------+
       ...          ...
      +------------+------------+------------+------------+------------+------------+
   20 |    19      | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+------------+

   [CLASS]

  And some data to insert into the workbook

     values( 20,'Bill','X',15,86.5,132)    * 20 is primary key UID in table above;
     values( 21,'Val', 'I',15,86.5,132)    * 1 to 1 with table layout;

WANT (You can only append to excel, sql has no knowledge of row order)


   d:/xls/class.xlsx

      +-----------------------------------------------------------------------------+
      |     A      |     B      |    C       |     D      |    E       |    F       |
      +-----------------------------------------------------------------------------+
   1  |   UID      | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+------------+
   2  |    1       | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+------------+
       ...          ...
      +------------+------------+------------+------------+------------+------------+
   20 |    19      | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+------------+
   21 |    20      | BILL       |    X       |    15      |   86.5     |  132       |   Added these
      +------------+------------+------------+------------+------------+------------+
   22 |    21      | VAL        |    I       |    15      |   86.5     |  132       |
      +------------+------------+------------+------------+------------+------------+

   [CLASS]


*                _               _       _
 _ __ ___   __ _| | _____     __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

;

 SAS dataset CLASS and EXCEL workbook

    %utlfkil(d:/xls/class.xlsx);

    libname xls  "d:/xls/class.xlsx" scan_text=no;

    data xls.class class;
        retain uid;
        set sashelp.class;
        uid=_n_;
    run;

    libname xls clear;


    proc sql dquote=ansi;
     connect to excel (Path="d:/xls/class.xlsx");
        execute(insert into class values( 20,'Bill','X',15,86.5,132)) by excel;
        execute(insert into class values( 21,'Val', 'I',15,86.5,132)) by excel;
    disconnect from Excel;
    quit;

    NOTE: There were 19 observations read from the data set SASHELP.CLASS.
    NOTE: The data set XLS.class has 19 observations and 6 variables.
    NOTE: The data set WORK.CLASS has 19 observations and 6 variables.

*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __
/ __|/ _ \| | | | | __| |/ _ \| '_ \
\__ \ (_) | | |_| | |_| | (_) | | | |
|___/\___/|_|\__,_|\__|_|\___/|_| |_|

;

filename tagset http "http://support.sas.com/rnd/base/ods/odsmarkup/sql.sas";
%include tagset;
ods tagsets.sql file="d:/sql/class.sql";
proc print data=class ;
run;
ods _all_ close;
ods listing;

This is what you get

Create table class(uid float, name varchar(), sex varchar(), age float, height float, weight float);
Insert into class(uid, name, sex, age, height, weight) Values(1, 'Alfred', 'M', 14, 69.0, 112.5);
Insert into class(uid, name, sex, age, height, weight) Values(2, 'Alice', 'F', 13, 56.5, 84.0);
Insert into class(uid, name, sex, age, height, weight) Values(3, 'Barbara', 'F', 13, 65.3, 98.0);
Insert into class(uid, name, sex, age, height, weight) Values(4, 'Carol', 'F', 14, 62.8, 102.5);
Insert into class(uid, name, sex, age, height, weight) Values(5, 'Henry', 'M', 14, 63.5, 102.5);
Insert into class(uid, name, sex, age, height, weight) Values(6, 'James', 'M', 12, 57.3, 83.0);
Insert into class(uid, name, sex, age, height, weight) Values(7, 'Jane', 'F', 12, 59.8, 84.5);
Insert into class(uid, name, sex, age, height, weight) Values(8, 'Janet', 'F', 15, 62.5, 112.5);
Insert into class(uid, name, sex, age, height, weight) Values(9, 'Jeffrey', 'M', 13, 62.5, 84.0);
Insert into class(uid, name, sex, age, height, weight) Values(10, 'John', 'M', 12, 59.0, 99.5);
Insert into class(uid, name, sex, age, height, weight) Values(11, 'Joyce', 'F', 11, 51.3, 50.5);
Insert into class(uid, name, sex, age, height, weight) Values(12, 'Judy', 'F', 14, 64.3, 90.0);
Insert into class(uid, name, sex, age, height, weight) Values(13, 'Louise', 'F', 12, 56.3, 77.0);
Insert into class(uid, name, sex, age, height, weight) Values(14, 'Mary', 'F', 15, 66.5, 112.0);
Insert into class(uid, name, sex, age, height, weight) Values(15, 'Philip', 'M', 16, 72.0, 150.0);
Insert into class(uid, name, sex, age, height, weight) Values(16, 'Robert', 'M', 12, 64.8, 128.0);
Insert into class(uid, name, sex, age, height, weight) Values(17, 'Ronald', 'M', 15, 67.0, 133.0);
Insert into class(uid, name, sex, age, height, weight) Values(18, 'Thomas', 'M', 11, 57.5, 85.0);
Insert into class(uid, name, sex, age, height, weight) Values(19, 'William', 'M', 15, 66.5, 112.0);


Edit the above

proc sql;
create table class(uid float, name varchar(8), sex varchar(1),
  age float, height float, weight float);
Insert into class(uid, name, sex, age, height, weight)
Values(1, 'Alfred', 'M', 14, 69.0, 112.5)
Values(2, 'Alice', 'F', 13, 56.5, 84.0)
Values(3, 'Barbara', 'F', 13, 65.3, 98.0)
Values(4, 'Carol', 'F', 14, 62.8, 102.5)
Values(5, 'Henry', 'M', 14, 63.5, 102.5)
Values(6, 'James', 'M', 12, 57.3, 83.0)
Values(7, 'Jane', 'F', 12, 59.8, 84.5)
Values(8, 'Janet', 'F', 15, 62.5, 112.5)
Values(9, 'Jeffrey', 'M', 13, 62.5, 84.0)
Values(10, 'John', 'M', 12, 59.0, 99.5)
Values(11, 'Joyce', 'F', 11, 51.3, 50.5)
Values(12, 'Judy', 'F', 14, 64.3, 90.0)
Values(13, 'Louise', 'F', 12, 56.3, 77.0)
Values(14, 'Mary', 'F', 15, 66.5, 112.0)
Values(15, 'Philip', 'M', 16, 72.0, 150.0)
Values(16, 'Robert', 'M', 12, 64.8, 128.0)
Values(17, 'Ronald', 'M', 15, 67.0, 133.0)
Values(18, 'Thomas', 'M', 11, 57.5, 85.0)
Values(19, 'William', 'M', 15, 66.5, 112.0)
;quit;

