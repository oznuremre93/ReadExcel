package ReadExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String args[]) throws IOException, InvalidFormatException {

		// File file = new File("ornek.xlsx");
		FileInputStream fis = new FileInputStream(new File("ornek.xlsx"));

		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		FormulaEvaluator forlulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
		for (Row row : sheet) {
			for (Cell cell : row) {
				switch (forlulaEvaluator.evaluateInCell(cell).getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.println(cell.getNumericCellValue() + "\t\t");

					break;
				case Cell.CELL_TYPE_STRING:
					System.out.println(cell.getStringCellValue() + "\t\t");

					break;

				default:
					break;
				}
			}
		}
		System.out.println();

	}

}
