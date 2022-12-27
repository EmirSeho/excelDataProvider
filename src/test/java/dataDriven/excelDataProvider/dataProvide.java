package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.swing.text.DateFormatter;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataProvide {
	
	DataFormatter formatter = new DataFormatter();
	
	@Test(dataProvider = "driveTest")
	public void testCaseData(String greeting, String communication, String id)
	{
		System.out.println(greeting + communication + id);
	}

	@DataProvider(name = "driveTest")
	public Object[][] getData() throws IOException
	{
		//Object[][] data = {{"hello", "test", "1"}, {"hello2", "test2", "12"}, {"hello3", "test3", "13"}};
		
		FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "//30.02excelDriven.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int columnCount = row.getLastCellNum();
		
		Object data[][] = new Object[rowCount - 1][columnCount];
		
		for(int i=0; i<rowCount - 1; i++)
		{
			row = sheet.getRow(i+1);
			for(int j=0; j<columnCount; j++)
			{
				XSSFCell cell = row.getCell(j);
				data[i][j] = formatter.formatCellValue(cell);
			}
		}
		
		return data;
	}

}
