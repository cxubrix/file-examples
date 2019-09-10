package lv.accenture.start.it;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// based on https://www.baeldung.com/java-microsoft-excel
public class ExcelRead {

	public static void main(String[] args) throws IOException {

		String filename = "hitmans.xlsx";
		File file = new File(filename);
		System.out.println("Usinng file: '" + file.getAbsoluteFile().toPath() + "'");

		ExcelRead er = new ExcelRead();
		Hitman[] hitmans = er.readExcel(file);
		
		// print array contents 
		for (int i = 0; i < hitmans.length; i++) {
			System.out.println(hitmans[i].getFirstname() + " " + hitmans[i].getLastname()+ " " + hitmans[i].getPrice());
		}
	}

	public Hitman[] readExcel(File file) throws IOException {

		FileInputStream fis = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheetAt(0);

		int lastRow = sheet.getLastRowNum();

		Hitman[] hitmans = new Hitman[lastRow]; // store all in array

		for (int i = 0; i < lastRow; i++) { // read each row
			Row row = sheet.getRow(i + 1); // shift by one, because of header
			String name = row.getCell(0).getRichStringCellValue().getString(); // reading string
			String surname = row.getCell(1).getRichStringCellValue().getString(); // reading string
			double price = row.getCell(2).getNumericCellValue(); // reading number

			// print each line
//			System.out.println(name + " " + surname + " " + price);

			// create object and store
			Hitman hitman = new Hitman(name, surname, price);
			hitmans[i] = hitman;

		}
		workbook.close();
		return hitmans;
	}

}
