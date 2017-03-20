package testngDataProvExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.File;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelDataConfig {
	XSSFWorkbook wb;
	XSSFSheet sheet;

	public excelDataConfig(String filePath) throws IOException {
		FileInputStream fis = new FileInputStream(
				"C:\\Selenium\\testngDataProviExcel\\src\\testngDataProvExcel\\excelData.xlsx");
		wb = new XSSFWorkbook(fis);
	}

	public String getData(int sheetNumber, int row, int column) {
		sheet = wb.getSheetAt(sheetNumber);
		String data = sheet.getRow(row).getCell(column).getStringCellValue();
		return data;
	}

	public int getRowCount(int sheetIndex) {
		int row = wb.getSheetAt(sheetIndex).getLastRowNum();
		row = row + 1;
		return row;
	}
}
