package excelDriven;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {
	
	//Identify Testcases coloum by scanning the entire 1st row
	//once coloumn is identified then scan entire testcase coloum to identify purchase testcase row
	//after you grab purchase testcase row = pull all the data of that row and feed into test
	
	public ArrayList<String> getData(String testCaseName) throws IOException 
	{
		ArrayList<String> aList = new ArrayList<String>();
		
		FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "//30.01DemoData.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheets = workbook.getNumberOfSheets();
		
		for (int i=0; i<sheets; i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				//Identify Testcases coloum by scanning the entire 1st row
				Iterator<Row> rows = sheet.iterator();
				Row firstRow = rows.next();//go to first row
				
				//go through cells in row
				Iterator<Cell> cells = firstRow.cellIterator();
				
				int k = 0;
				int column = 0;
				while(cells.hasNext())
				{
					Cell value = cells.next();//go to first cell in row
					if(value.getStringCellValue().equalsIgnoreCase("testcases"))
					{
						column = k;
					}
					
					k++;
				}
				
				System.out.println(column);
				
				//once coloumn is identified then scan entire testcase coloum to identify purchase testcase row
				while(rows.hasNext())
				{
					Row r = rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName))
					{
						//after you grab purchase testcase row = pull all the data of that row and feed into test
						Iterator<Cell> cv = r.cellIterator();
						while(cv.hasNext())
						{
							Cell c = cv.next();
							if(c.getCellType() == CellType.STRING)
							{
								aList.add(c.getStringCellValue());
							}
							else
							{
								aList.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
							
						}
					}
					
				}
				
			}
			
	}
		return aList;
		
	}
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub		
	}
}
