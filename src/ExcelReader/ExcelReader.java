package ExcelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Cell;
//
//import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	public static void main(String[] args) {
		try {
			//  Replace your excel url with the following url (attempted to use googleSheet file)
			// String fileUrlStr =
			// "https://docs.google.com/spreadsheets/d/e/2PACX-1vQDhukToZsrxdHqr6zbbOmAyLCmz8sc6yeuuCRx3DnB4IV_08OP1LgPgTlxZFEeV3nDzv7ajoEgO-8G/pubhtml";
			String fileLocalStr = "src/Test.xlsx";

			FileInputStream file = new FileInputStream(new File(fileLocalStr));

			// uncomment this line when using URL.
			// InputStream fileUrl = new URL(fileUrlStr).openStream();

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(fileLocalStr);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					switch (cell.getCellType()) {
					case NUMERIC:
						if (HSSFDateUtil.isCellDateFormatted(cell)) {
							System.out.print(cell.getDateCellValue() + "\t");
						} else {
							System.out.print(cell.getNumericCellValue() + "\t");
						}
						break;
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					}
				}
				System.out.println("");
			}
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
