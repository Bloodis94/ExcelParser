# ExcelParser
My first attempt at writing a library that wraps Apache POI to speed up usage.

The ExcelParserExecutor class expects an implementation of XSSFWorkbook and Destination to work.

The ExcelParserAbstract class can be extended to customize the parsing process.

The library uses reflection to read the first row of an excel file, and to map the names of each column to the homonymous variables in a class that implements Destination, returning an instance of the implementation for each row of the excel file
