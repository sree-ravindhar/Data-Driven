package Data.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\ravi2\\OneDrive\\Documents\\Test.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheetAt = wb.getSheetAt(0);
		int rows = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < rows; i++) {

			XSSFRow row = sheetAt.getRow(i);
			int cellCount = row.getPhysicalNumberOfCells();
			for (int j = 0; j < cellCount; j++) {

				XSSFCell cell = row.getCell(j);
				CellType cellType = cell.getCellType();

				if (cellType.equals(CellType.STRING)) {

					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);

				}
				if (cellType.equals(cellType.NUMERIC)) {

					double numericCellValue = cell.getNumericCellValue();
					// System.out.println(numericCellValue);
					long s = (long) numericCellValue;
					System.out.println(s);
					String valueOf = String.valueOf(s);
					System.out.println(valueOf);

				}

			}

		}
		System.out.println("------Print exact value for the given location");
		String stringCellValue = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
		System.out.println(stringCellValue);

		XSSFCell createCell = wb.getSheetAt(1).createRow(0).createCell(0);
		createCell.setCellValue("ravi230998@gmail.com");
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		wb.close();
		System.out.println("Value Written");

	}

}
