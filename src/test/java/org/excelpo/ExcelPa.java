package org.excelpo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPa {

	public static void excelMetho() throws IOException {
		// TODO Auto-generated method stub
		//TODO Auto-generated method stub6
		File f = new File("C:\\Users\\Admin\\eclipse-workspace\\ExcelPAth\\target\\test1.xlsx");
		FileInputStream stream = new FileInputStream(f);
		XSSFWorkbook wo= new XSSFWorkbook(stream);
		XSSFSheet sheet = wo.getSheet("data");
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			XSSFRow row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				 XSSFCell cell = row.getCell(j);
				 int cellType = cell.getCellType();
				if (cellType==1) {
				String stringCellValue = cell.getStringCellValue();
				if (stringCellValue.equals("shahu")) {
					String stringCellValue2 = row.getCell(1).getStringCellValue();
					XSSFCell cell2 = row.getCell(1);
					cell2.setCellValue("shahul hameed h");
					FileOutputStream fO = new FileOutputStream(f);
					wo.write(fO);
				}
				}
			}
			
		}
			}
	
	public static void main(String[] args) throws IOException {
		excelMetho();
	}
}
