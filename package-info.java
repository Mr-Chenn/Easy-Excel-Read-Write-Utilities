/**
 * 
 */
/**
 * @author Mr-Chenn
 * This package has two classes namely
 * 1. ExcelReadClass 
 * 		(a) This class has very easy to use methods. To read from a specified row, column of an excel
 * 		(b) Also you can ready based on column header name when you can't rely on the column index
 * 		(c) Also a method to get the row & column count from an excel sheet
 * 2. ExcelWriteClass
 * 		(a) This class has a very easy to use write methods. To write to a specified row, column of excel. Actual row number, column number, not 0 based index
 * 		(b) You can also specify the column header name and the row number within the column to write.
 * 
 * EASY EXAMPLE TO MAKE CALLS TO THE ABOVE METHODS -->
 *		write.setDataToExcelCell(23, 3,"HELLO WORLD 23-3."); //Writes to 23rd excel row, 3rd excel column(C) 
 *		System.out.println("Reading the written value =====\n" +read.getDataFromExcelCell(22, 2)); //prints above value
 *		write.setDataToExcelCell("column 8", 10, "HELLO beautiful WORLD COLUMN8-10"); //Writes to 10th cell value under column labeled "column 8"
 *		System.out.println("Reading the written value =====\n" +read.getDataFromExcelCell("column 8", 10)); //read above value
 * 		
 * mORE Examples of making calls to these methods-->
 * Get Row count or column count
 * 		ExcelReadClass read = new ExcelReadClass();
 *		int numOfColumns = read.getRowColumnCount("Columns");
 *		int numOfRows = read.getRowColumnCount("Rows");
 * Read the value from 4th row and 5th column
 * 		System.out.println("***\n"+read.getDataFromExcelCell(3, 4));
 * Read the value from 5th cell value (0 refers to column header) in column labeled as "Column6"
 * 		System.out.println(read.getDataFromExcelCell("Column6", 5));
 * 
 * Write to 20th row 17th column in excel. NOT index 0 based. count rownum & colnum startig 1.
 * 	 	ExcelWriteClass write = new ExcelWriteClass();
 *		write.setDataToExcelCell(20, 17,"HELLO WORLD 20-17.");
 *		System.out.println("Reading the written value =====\n" +read.getDataFromExcelCell(19, 16)); //To read above set value
 * Write to 5th cell(0-->header) in the column labeled as "Column7" 		
 * 		write.setDataToExcelCell("Column7", 5, "HELLO WORLD COLUMN7-5");
 *		System.out.println("Reading the written value =====\n" +read.getDataFromExcelCell("Column7", 5));  
 *
 * 
 * 
 * 
 */
package excelUtils;