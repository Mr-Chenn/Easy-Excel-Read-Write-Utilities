package excelUtils;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Properties;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadClass 
{
	public static String propertiesFilePath = "C:\\selenium\\Global properties file\\GlobalFile.properties";
	
	/***
	 * Below method returns the count of rows or columns in the excel sheet. NOT 0 based. Actual counts.
	 * @param arg --> takes "Rows" or "Columns" as an argument 
	 * @return --> returns the number of rows or columns based on the argument passed. 
	 * @throws IOException
	 */
	public int getRowColumnCount(String arg) throws IOException 
	{	
		arg = arg.toUpperCase();
		
		//Loading the properties file
		Properties prop = new Properties();
		FileInputStream fisProp = new FileInputStream(propertiesFilePath);
		prop.load(fisProp);
		
		FileInputStream fis = new FileInputStream(prop.getProperty("EXCELFILEPATH"));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet(prop.getProperty("SHEETNAME"));
				
		int numOfRows = sheet.getLastRowNum() + 1;
		int numOfColumnHeaders = sheet.getRow(0).getLastCellNum();
		
		if(arg.equals("ROWS"))
			return numOfRows;
		else if(arg.equals("COLUMNS"))
			return numOfColumnHeaders;
		else
		{
			System.out.println("Invalid argument");
			return 0;
		}	
		 
	}
		
	/***
	 * Below method returns the value from an excel cell (rownum, colnum) --> 0 based index
	 * @param rownum --> 0,1,2,3....
	 * @param colnum --> 0,1,2,3....
	 * @return
	 * @throws IOException
	 */
	@SuppressWarnings("finally")
	public String getDataFromExcelCell(int rownum, int colnum) throws IOException
	{
		String cellValue="";
		
		Properties prop = new Properties();
		FileInputStream fisProp = new FileInputStream(propertiesFilePath);
		prop.load(fisProp);
		
		FileInputStream fis = new FileInputStream(prop.getProperty("EXCELFILEPATH"));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet(prop.getProperty("SHEETNAME"));
		XSSFRow row = sheet.getRow(rownum);
		XSSFCell cell = row.getCell(colnum);
	
		try {
			cellValue = cell.getStringCellValue();
		} catch (Exception e) {
			cellValue = null;
		}		
		finally {		
			return cellValue;
		}
		
	}
	
	/***
	 * Below method returns the index of a column header. 0 based index.
	 * This method is declared as with no access modifier. so it can be accessed only within the package
	 * Make sure there are no duplicate column names in the excel file. This program can index unique column names
	 * @param columnHeaderName --> column header name (string) --> Note this is case sensitive. You can easily add logic to remove case-sensitivity
	 * @return  --> returns the index of the column header. 0, 1, 2, 3 etc...    
	 * 			--> returns -1 if the column name passed is not available.          
	 * @throws IOException
	 */
	int getIndexOfExcelColumn(String columnHeaderName) throws IOException
	{
		 
		 int colIndex=0;
		 LinkedHashMap<String,Integer> lhMap = new LinkedHashMap<String,Integer>();
		 int numOfColumnHeaders = getRowColumnCount("Columns");
		 
		 for(int i=0;i<numOfColumnHeaders;i++)
		 {
			 lhMap.put(getDataFromExcelCell(0, i), i);
		 }
		 System.out.println(lhMap);		 
		 
		 try {
			 colIndex = lhMap.get(columnHeaderName);		 
		 } catch (Exception e) {
			 System.out.println("Pass a valid column name");
			 colIndex = -1;
		 }
		 return colIndex;
	}
	
	/***
	 * Below method returns the value of a cell based on a column header name. 0 based index.
	 * Make sure there are no duplicate column names in the excel file. 
	 * @param columnHeaderName --> name of the column header --> Is case sensitive.
	 * @param rownum --> pass 5 to get the value from 5th cell under the column header (same as from excel row 6)
	 * 				 --> 0 refers to header name, 1 refers to 1st cell value under the column header.....4 refers to 4th cell value under the column header
	 * @return  --> value of the cell
	 * @throws IOException
	 * 
	 * More logic can be added if the column name is not available. i.e., can add validation steps or if-else logic if the colnum = -1;
	 */
	public String getDataFromExcelCell(String columnHeaderName, int rownum) throws IOException
	{	
		int colnum = getIndexOfExcelColumn(columnHeaderName);
		//System.out.println("Index of "+columnHeaderName +" is = "+colnum);
		String cellValue = getDataFromExcelCell(rownum,colnum);
		return cellValue;		
	}
	
	
}

