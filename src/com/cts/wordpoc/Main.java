package com.cts.wordpoc;

import java.io.IOException;
import java.time.LocalDate;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import com.cts.wordpoc.service.WordDocumentUtil;

public class Main {

	public static void main(String[] args) throws IOException, XmlException {
		// TODO Auto-generated method stub
		
		System.out.println("start");
		 
		
		WordDocumentUtil wordUtil = new WordDocumentUtil(); 
		XWPFDocument doc = wordUtil.create(); 
		
		wordUtil.addTitle("My Title shanu", 16, ParagraphAlignment.CENTER); 
		wordUtil.addParagraph("This is my default size paragraph"); 
		wordUtil.addParagraph("This is my large font size paragraph", 14);
		
		
		try {
			wordUtil.addHeader("This is my header");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		try {
			wordUtil.addFooter("This is my footer");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		
		
		 LocalDate date=LocalDate.now();
		 wordUtil.save("my_test_doc_"+date);
		
		 System.out.println("end");


	}

}
