package com.datadriven;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDrivenn {
	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\Vinothraj\\Desktop\\DataBook.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);
		int rowCount = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < rowCount; i++) {
			Row row = sheetAt.getRow(i);
			int cellCount = row.getPhysicalNumberOfCells();
			for (int j = 0; j < cellCount; j++) {
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				if (cellType.equals(cellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
				}
				if(cellType.equals(cellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					long l = (long) numericCellValue;
					String valueOf = String.valueOf(l);
					System.out.println(valueOf);
				}
			}
		}
		System.out.println("------Print exact value for the given location");
		String resultExact = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
		System.out.println(resultExact);
		double numericCellValue1 = wb.getSheetAt(0).getRow(0).getCell(1).getNumericCellValue();
		long l1 = (long) numericCellValue1;
		String v = String.valueOf(l1);
		System.out.println(v);
	
		System.out.println("------------------create the cell and set the string value-------");
		Cell createCell = wb.getSheetAt(1).createRow(5).createCell(0);
		createCell.setCellValue("Vinothraj");
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		wb.close();
		System.out.println("Value inserted");
		
	}
}