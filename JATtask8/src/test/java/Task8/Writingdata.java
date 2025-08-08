package Task8;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;


public class Writingdata {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String excelFilePath=".\\datafile\\write.xlsx";
		FileInputStream inputstream=new FileInputStream(excelFilePath);
		XSSFWorkbook workbook= new XSSFWorkbook(inputstream);
		XSSFSheet sheet=workbook.getSheetAt(0);
		
		Object citycode[][] = { {"City", "State", "Populationinlakh"},
				{"Chennai", "Tamilnadu", 20},
				{"Bangalore", "Karnataka", 30},
				{"Hyderabad", "Telengana", 15},
				{"Mumbai", "Maharashtra", 40}

		 };

				int rows=citycode.length;
				int cols=citycode[0].length;
				System.out.println(rows);
				System.out.println(cols);
				for(int r=0;r<rows;r++) {
					XSSFRow row=sheet.createRow(r);
					for(int c=0;c<cols;c++) {
						XSSFCell cell= row.createCell(c);
						Object value=citycode[r][c];
						
						if(value instanceof String)
							cell.setCellValue((String)value);
						if(value instanceof Integer)
							cell.setCellValue((Integer)value);
						if(value instanceof Boolean)
							cell.setCellValue((Boolean)value);
					}
					
				}
				FileOutputStream outstream=new FileOutputStream(excelFilePath);
				workbook.write(outstream);
				outstream.close();
				System.out.println("File is writed successfully");
		}
		}