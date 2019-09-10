package lv.accenture.start.it;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// based on https://www.baeldung.com/java-microsoft-excel
public class ExcelWrite {

	public static void main(String[] args) throws IOException {
		String filename = "output.xlsx";
		File file = new File(filename);

		System.out.println("Usinng file: '" + file.getAbsoluteFile().toPath() + "'");

		ExcelWrite ew = new ExcelWrite();
		ew.writeExcel(file, generate(10));
	}

	public void writeExcel(File file, Hitman[] hitmans) throws IOException {
		Workbook workbook = new XSSFWorkbook(); // create workbook
		Sheet sheet = workbook.createSheet(); // create sheed

		XSSFFont font = ((XSSFWorkbook) workbook).createFont();
		font.setBold(true);

		CellStyle headerStyle = workbook.createCellStyle(); // style for header
		headerStyle.setFont(font); // assign font to header style

		Row header = sheet.createRow(0); // create header

		Cell firstnameHeader = header.createCell(0); // create header cell for fistname
		firstnameHeader.setCellStyle(headerStyle);
		firstnameHeader.setCellValue("Name");

		Cell lastnameHeader = header.createCell(1); // create header cell for lastname
		lastnameHeader.setCellStyle(headerStyle);
		lastnameHeader.setCellValue("Surname");

		Cell priceHeader = header.createCell(2); // create header cell for price
		priceHeader.setCellStyle(headerStyle);
		priceHeader.setCellValue("Price");

		for (int i = 0; i < hitmans.length; i++) {
			Row row = sheet.createRow(i + 1); // mind the header row

			Cell fistnameCell = row.createCell(0); // firstname value cell
			fistnameCell.setCellValue(hitmans[i].getFirstname());

			Cell lastnameCell = row.createCell(1); // lastname value cell
			lastnameCell.setCellValue(hitmans[i].getLastname());

			Cell priceCell = row.createCell(2); // price value cell
			priceCell.setCellValue(hitmans[i].getPrice());
		}

		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		workbook.close();

		System.out.println("Write");

	}

	private static Hitman[] generate(int num) {
		Hitman[] hitmans = new Hitman[num];
		for (int i = 0; i < num; i++) {
			double price = 3 * i + (i % 4 / (double) 0.97); // just to look random
			hitmans[i] = new Hitman("Firstname_" + i, "Lastname_" + i, price); // set value in array
		}
		return hitmans;
	}
}
