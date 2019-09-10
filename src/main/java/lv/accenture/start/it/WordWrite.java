package lv.accenture.start.it;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

// based on https://www.baeldung.com/java-microsoft-word-with-apache-poi
public class WordWrite {

	public static void main(String[] args) throws FileNotFoundException, IOException {

		// dummy value to store in doc
		Hitman hitman = new Hitman("David", "Bateson", 1_000_000.9999);

		File templateFile = new File("template.docx"); // template file
		File outputFile = new File("output.docx"); // output file

		WordWrite ww = new WordWrite();
		ww.writeWord(templateFile, outputFile, hitman);
	}

	public void writeWord(File templateFile, File outputFile, Hitman hitman)
			throws FileNotFoundException, IOException {

		XWPFDocument document = new XWPFDocument(new FileInputStream(templateFile));
		List<XWPFParagraph> paragraphs = document.getParagraphs();

//		XWPFParagraph title = paragraphs.get(0);
//		System.out.println(title.getText());
//		System.out.println(paragraphs.size());

		XWPFParagraph content = paragraphs.get(1); // next one after title

		// read text and modify
		String text = content.getText();
		text = text.replace("{firstname}", hitman.getFirstname());
		text = text.replace("{lastname}", hitman.getLastname().toUpperCase());
		text = text.replace("{price}", "$" + hitman.getPrice());

		// set back content to modified text
		content.getRuns().get(0).setText(text, 0);

		// save to output file
		FileOutputStream out = new FileOutputStream(outputFile);
		document.write(out);
		out.close();

		document.close();
	}
}
