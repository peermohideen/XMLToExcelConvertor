# XMLToExcelConvertor
TO CONVERT XML FILE TO ExCEL formaT

This is simplest code to Read The XML file and Wirte into MSExcel format.

IF We are using without the IDE we can use the following command

Compile commmand:
javac -cp C:\Users\730024\POC\jars\poi-3.17.jar;poi-ooxml-3.17.jar ExcelCon.java

Executing command:
java -cp C:\Users\730024\POC\jars\poi-3.17.jar;poi-ooxml-3.17.jar; ExcelCon




NOTE:
To read XLS files, an HSSF implementation is provided by POI library.

To read XLSX, XSSF implementation of POI library.



XLS classes         Interfaces          XLSXCLASSES
===================================================
HSSFWorkbook        Workbook            XSSFWorkbook
HSSFSheet           Sheet               XSSFSheet
HSSFRow             Row                 XSSFRow
HSSFCell            Cell                XSSFCell




Workbook: XSSFWorkbook and HSSFWorkbook classes implement this interface.

XSSFWorkbook: Is a class representation of XLSX file.

HSSFWorkbook: Is a class representation of XLS file.

Sheet: XSSFSheet and HSSFSheet classes implement this interface.

XSSFSheet: Is a class representing a sheet in an XLSX file.

HSSFSheet: Is a class representing a sheet in an XLS file.

Row: XSSFRow and HSSFRow classes implement this interface.

XSSFRow: Is a class representing a row in the sheet of XLSX file.

HSSFRow: Is a class representing a row in the sheet of XLS file.

Cell: XSSFCell and HSSFCell classes implement this interface.

XSSFCell: Is a class representing a cell in a row of XLSX file.

HSSFCell: Is a class representing a cell in a row of XLS file.
