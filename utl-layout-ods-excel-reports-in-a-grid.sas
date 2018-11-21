%let pgm=utl-layout-ods-excel-reports-in-a-grid;

Layout ods excel reports in a grid (windows and unix)

Output Excel (side by side reports)
https://tinyurl.com/y7g552rt
https://github.com/rogerjdeangelis/utl-layout-ods-excel-reports-in-a-grid/blob/master/sex_after_layout.xlsx

github
https://github.com/rogerjdeangelis/utl-layout-ods-excel-reports-in-a-grid

Source
https://yagisanatode.com/2017/11/18/copy-and-paste-ranges-in-excel-with-openpyxl-and-python-3/

There appears to be an issue copying titles and footnotes?

INPUT
=====

   d:/xls/sex.xlsx

   Two 'proc report sheets' [sex] and [sex2]

      MALES

      +--------------------------------------+
      |     A      |    B       |     C      |
      +--------------------------------------+
   1  | NAME       |   SEX      |    AGE
      +------------+------------+------------+
   2  | ALFRED     |    M       |    14      |
      +------------+------------+------------+
       ...
      +------------+------------+------------+
   13 | WILLIAM    |    M       |    15      |
      +------------+------------+------------+

   [SEX]

      FEMALES

      +--------------------------------------+
      |     A      |    B       |     C      |
      +--------------------------------------+
   1  | NAME       |   SEX      |    AGE
      +------------+------------+------------+
   2  | ALICE      |    M       |    14      |
      +------------+------------+------------+
       ...
      +------------+------------+------------+
   11 | SUSAN      |    M       |    15      |
      +------------+------------+------------+

   [SEX2]


WANT
====

      MALES                                      FEMALES

      +--------------------------------------+   +--------------------------------------+
      |     A      |    B       |     C      |   |     E      |    F       |     G      |
      +--------------------------------------+   +--------------------------------------+
   1  | NAME       |   SEX      |    AGE         | NAME       |   SEX      |    AGE
      +------------+------------+------------+   +------------+------------+------------+
   2  | ALFRED     |    M       |    14      |   | ALICE      |    M       |    14      |
      +------------+------------+------------+   +------------+------------+------------+
       ...                                        ...
      +------------+------------+------------+   +------------+------------+------------+
  13  | WILLIAM    |    M       |    15      |   |            |            |            |
      +------------+------------+------------+   +------------+------------+------------+


PROCESS
=======

%utl_submit_py64('
import openpyxl;
wb = openpyxl.load_workbook("d:/xls/sex.xlsx");
sheet = wb["sex2"];
temp_sheet = wb["sex"];
def copyRange(startCol, startRow, endCol, endRow, sheet):;
.   rangeSelected = [];
.   for i in range(startRow,endRow + 1,1):;
.       rowSelected = [];
.       for j in range(startCol,endCol+1,1):;
.           rowSelected.append(sheet.cell(row = i, column = j).value);
.       rangeSelected.append(rowSelected);
.   return rangeSelected;
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData):;
.   countRow = 0;
.   for i in range(startRow,endRow+1,1):;
.       countCol = 0;
.       for j in range(startCol,endCol+1,1):;
.           sheetReceiving.cell(row = i, column = j).value = copiedData[countRow][countCol];
.           countCol += 1;
.       countRow += 1;
def createData():;
.   print("Processing...");
.   selectedRange = copyRange(1,1,3,13,sheet);
.   pastingRange = pasteRange(5,1,7,11,temp_sheet,selectedRange);
.   wb.remove(sheet);
.   wb.save("d:/xls/sex.xlsx");
.   print("Range copied and pasted!");
createData();
');

*                _               _       _
 _ __ ___   __ _| | _____     __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

;

%utlfkil(d:/xls/sex.xlsx);

title;footnote;
ods excel file="d:/xls/sex.xlsx" style=minimal
  Options( sheet_name="sex2");

proc report data=sashelp.class(
     keep= name sex age
     where=(sex='M')) missing nowd;
run;quit;

ods excel Options( sheet_interval="none" sheet_name="sex");
proc report data=sashelp.class(
     keep= name sex age
     where=(sex='F')) missing nowd;
run;quit;

ods excel close;

