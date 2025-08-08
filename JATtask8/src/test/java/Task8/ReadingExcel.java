package Task8;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.*;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String excelFilePath=".\\datafile\\test1.xlsx";
		FileInputStream inputstream=new FileInputStream(excelFilePath);
		XSSFWorkbook workbook= new XSSFWorkbook(inputstream);
		XSSFSheet sheet=workbook.getSheetAt(0);
		int rows= sheet.getLastRowNum();
		int cols=sheet.getRow(1).getLastCellNum();
		System.out.println(cols);
		System.out.println(rows);
		
		for(int r=0;r<=rows;r++) {
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<cols;c++) {
				XSSFCell cell=row.getCell(c);
				switch(cell.getCellType()) {
				case NUMERIC : System.out.print(cell.getNumericCellValue());break;
				case STRING : System.out.print(cell.getStringCellValue());break;			
				case BOOLEAN : System.out.print(cell.getBooleanCellValue());break;
				
				}
				System.out.print("| ");
				
			}
			System.out.println();
		}
		
		

	}

}
