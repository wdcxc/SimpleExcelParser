# SimpleExcelParser

a simple excel data parser for php base on PHPExcel.

you can read excel data as an array or write the data into a file,you can get some simple examples in test.php

if the excel data is large, it's better to modify the code in SimpleExcelParser.php, such as the var chunkSize and the procedure in function getData() and extractDataIntoFile()
