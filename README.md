# utl-appending-records-to-an-existing-excel-sheet
Appending records to an existing excel sheet.
    Appending records to an existing excel sheet

    see github
    https://tinyurl.com/yalhj2o4
    https://github.com/rogerjdeangelis/utl-appending-records-to-an-existing-excel-sheet

    SAS  Forum
    https://communities.sas.com/t5/New-SAS-User/ODS-EXCEL/m-p/522637

    other excel repos
    https://tinyurl.com/ybnm6azh
    https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=


    INPUT
    =====

      d:/xls/class.xlsx

          +----------------------------------------------------------------+
          |     A      |    B       |     C      |    D       |    E       |
          +----------------------------------------------------------------+
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
          +------------+------------+------------+------------+------------+
       2  | ALFRED     |    M       |    14      |    69      |  112.5     |
          +------------+------------+------------+------------+------------+
       3  | ALICE      |    F       |    15      |   66.5     |  112       |
          +------------+------------+------------+------------+------------+

       [CLASS]

     * APPEND THESE;

     WORK.HAV1ST total obs=2

         NAME      SEX    AGE    HEIGHT    WEIGHT

        WWWWWWW     M      14     69.0      112.5
        XXXXXXX     F      13     56.5       84.0


     WORK.HAV2ND total obs=2

         NAME      SEX    AGE    HEIGHT    WEIGHT

        YYYYYYY     M      14     69.0      112.5
        ZZZZZZZ     F      13     56.5       84.0


    EXAMPLE OUTPUT
    --------------

      d:/xls/class.xlsx

          +----------------------------------------------------------------+
          |     A      |    B       |     C      |    D       |    E       |
          +----------------------------------------------------------------+
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
          +------------+------------+------------+------------+------------+
       2  | ALFRED     |    M       |    14      |    69      |  112.5     |
          +------------+------------+------------+------------+------------+
       3  | ALICE      |    F       |    15      |   66.5     |  112       |
          +------------+------------+------------+------------+------------+ *************************
       4  | WWWWWWW    |    M       |    15      |   66.5     |  112       |
          +------------+------------+------------+------------+------------+
       5  | XXXXXXX    |    F       |    15      |   66.5     |  112       |  APPEND
          +------------+------------+------------+------------+------------+  SAS TABLS HAV1ST HAV2ND
       6  | YYYYYYY    |    M       |    15      |   66.5     |  112       |
          +------------+------------+------------+------------+------------+
       7  | ZZZZZZZ    |    F       |    15      |   66.5     |  112       |
          +------------+------------+------------+------------+------------+ **************************

       [CLASS]


    PROCESS
    =======

    * Yeterdays append;
    libname xel "d:/xls/class.xlsx" scan_text=no;
    proc append base=xel.class data=hav1st;
    run;quit;
    libname xel clear;

    * Todays append;
    libname xel "d:/xls/class.xlsx" scan_text=no;
    proc append base=xel.class data=hav2nd;
    run;quit;
    libname xel clear;

    *                _               _       _
     _ __ ___   __ _| | _____     __| | __ _| |_ __ _
    | '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
    | | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
    |_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

    ;

    * create sheet;

    %utlfkil(d:/xls/class.xlsx);
    libname xel "d:/xls/class.xlsx";
    data xel.have;
      set sashelp.class(obs=2);
    run;quit;
    libname xel clear;

    data hav1st;
      set sashelp.class(obs=2);
      select (_n_);
         when (1) name="WWWWWWW";
         when (2) name="XXXXXXX";
      end;
    run;quit;

    data hav2nd;
      set sashelp.class(obs=2);
      select (_n_);
         when (1) name="YYYYYYY";
         when (2) name="ZZZZZZZ";
      end;
    run;quit;




