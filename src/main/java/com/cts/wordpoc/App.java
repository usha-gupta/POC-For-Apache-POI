package com.cts.wordpoc;

import java.io.IOException;
import java.time.LocalDate;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
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
        String[] tableHeaders = {"Name", "Role", "Company","Position"};
        String[][] tableData = { {"Hardware Kent", "Senior Network Architect", "NRT","Manager"}, 
        								{"Alex Adshead", "Senior Network Supporting Engineer", "","Manager"}, 
        								{"Amit Gautam",  "Releases Manager", "Cognizant ","Manager"}, 
        								{"Narendra Suraj", "project manager", "Cognizant","Manager"}};
        String[][] innerTableData = {{"Cognizant", "100"}, {"NRT", "100"},{"Cognizant", "100"}, {"NRT", "100"},{"Cognizant", "100"}};
        
        WordDocumentUtil wordUtil = new WordDocumentUtil(); 
		
			try {
				XWPFDocument doc = wordUtil.create();
				wordUtil.addTitle("My Title", 16, ParagraphAlignment.CENTER); 
				wordUtil.addParagraph("This is my default size paragraph"); 
				wordUtil.addParagraph("This is my large font size paragraph", 14);				
				LocalDate currentDate = LocalDate.now();
//				wordUtil.addHeader("This is my header....");
//				wordUtil.addFooter("This is my footer");				
				XWPFTable table = wordUtil.addTable(tableHeaders, "4d82be");
				wordUtil.addRows(table, tableData, "dbe5f1", "feffff");
				XWPFTableCell innerCell = table.getRow(2).getCell(2); 
				wordUtil.addNestedTable(innerCell,innerTableData);				
				wordUtil.save("my_test_doc_" + currentDate);
				
			} catch (IOException e ) {
				e.printStackTrace();
			}
				catch (XmlException e) {
				e.printStackTrace();
			} 	
    }
}
