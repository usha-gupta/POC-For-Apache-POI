package com.cts.wordpoc;

import java.io.IOException;
import java.time.LocalDate;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;

import com.cts.wordpoc.service.WordDocumentUtil;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
        
        
        WordDocumentUtil wordUtil = new WordDocumentUtil(); 
		
			try {
				XWPFDocument doc = wordUtil.create();
				wordUtil.addTitle("My Title", 16, ParagraphAlignment.CENTER); 
				wordUtil.addParagraph("This is my default size paragraph"); 
				wordUtil.addParagraph("This is my large font size paragraph", 14);
				
				LocalDate currentDate=LocalDate.now();
				wordUtil.addHeader("This is my header....");
				wordUtil.addFooter("This is my footer");
				wordUtil.save("my_test_doc_"+currentDate);
				
			} catch (IOException e ) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
				catch (XmlException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} 
			
    }
}
