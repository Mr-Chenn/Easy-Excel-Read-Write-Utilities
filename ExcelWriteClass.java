package excelUtils;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriteClass 
{
	public static String propertiesFilePath = "C:\\selenium\\Global properties file\\GlobalFile.properties";
	
	/***
	 * Below method returns the count of rows or columns in the excel sheet. NOT 0 based. Actual counts.
	 * This one calls the method from ExcelReadClass
	 * @param arg --> takes "Rows" or "Columns" as an argument 
	 * @return --> returns the number of rows or columns based on the argument passed. 
	 * @throws IOException
	 */
	public int getRowColumnCount(String arg) throws IOException 
	{	
		ExcelReadClass r = new ExcelReadClass();
		return r.getRowColumnCount(arg);		 
	}
		
	/***
	 * Write the value to an excel cell. rownum-th row, colnum-th column. NOT 0 based. 
	 * If the cell doesn't exist, this method will create blank row-cell & then writes to it
	 * @param rownum --> NOT 0 based index row number. Pass 9, to write to excel row 9.
	 * @param colnum --> NOT 0 based index column number. Pass 3, to write to excel Column - C.
	 * @param text --> text to be written to excel cell
	 * @throws IOException
	 */
	public void setDataToExcelCell(int rownum, int colnum, String text) throws IOException
	{
		rownum = rownum-1;
		colnum = colnum-1;
		Properties prop = new Properties();
		FileInputStream fisProp = new FileInputStream(propertiesFilePath);
		prop.load(fisProp);
		
		FileInputStream fis = new FileInputStream(prop.getProperty("EXCELFILEPATH"));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet(prop.getProperty("SHEETNAME"));
		
		try {
			XSSFRow row = sheet.getRow(rownum);
			XSSFCell cell = row.getCell(colnum);
			cell.setCellValue(text);
		}
		catch (Exception e) {
			System.out.println("The cell doesn't exist. Creating one.");
			sheet.createRow(rownum);
			XSSFRow row = sheet.getRow(rownum);
			row.createCell(colnum);
			XSSFCell cell = row.getCell(colnum);
			cell.setCellValue(text);
			
		}
		
		FileOutputStream fos = new FileOutputStream(prop.getProperty("EXCELFILEPATH"));
		wb.write(fos);
		fos.close();
		fis.close();
	}
	
	
	/***
	 * Below method writes a string text to an excel cell based on column header name & row number. Make sure column name exists. Column header is case sensitive
	 * @param columnHeaderName --> is the column header name in the excel in which you want to write a text. Is case sensitive. Logic to remove case-sensitivity can easily be added.
	 * @param rownum --> is the cell number in the row. 0 references to column header. pass 5 to set 5th cell in column(excel row =6). 
	 * @param text  --> is the string text you want to write to the excel
	 * @throws IOException
	 */
	public void setDataToExcelCell(String columnHeaderName, int rownum, String text) throws IOException
	{
		ExcelReadClass r = new ExcelReadClass();
		int colnum = r.getIndexOfExcelColumn(columnHeaderName);
		System.out.println("Index of "+columnHeaderName +" is = "+colnum);
		setDataToExcelCell(rownum + 1, colnum + 1, text); //adding 1 to rownum, colnum because - set method is not 0 based
	}
}

