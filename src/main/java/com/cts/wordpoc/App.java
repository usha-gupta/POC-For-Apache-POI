package com.cts.wordpoc;

import java.io.IOException;
import java.time.LocalDate;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlException;

import com.cts.wordpoc.service.WordDocumentUtil;

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
				
				String headingId  = "heading1";
				String subHeadingId = "heading2";
				
				wordUtil.addCustomHeadingStyle(headingId, 1);
				wordUtil.addCustomHeadingStyle(subHeadingId,2);
				
				wordUtil.addTitle("This is my title", 16, ParagraphAlignment.CENTER,headingId );				
				wordUtil.addParagraph("This is my paragraph for heading one");
				
				
				 wordUtil.addTitle("This is my second title--2",16,ParagraphAlignment.LEFT,headingId);
				 wordUtil.addParagraph("This is my paragraph for second for heading----2");
				 
				 
				wordUtil.addTitle("this is sub heading for heading 2", 14,ParagraphAlignment.LEFT,subHeadingId);
				wordUtil.addParagraph("This is my paragraph for subheading for paragph 2");
				
				
//				wordUtil.addParagraph("This is my default size paragraph"); 
//				wordUtil.addParagraph("This is my large font size paragraph", 14);				
				LocalDate currentDate = LocalDate.now();
//				wordUtil.addHeader("This is my header....");
//				wordUtil.addFooter("This is my footer");				
//				XWPFTable table = wordUtil.addTable(tableHeaders, "4d82be");
//				wordUtil.addRows(table, tableData, "dbe5f1", "feffff");
//				XWPFTableCell innerCell = table.getRow(2).getCell(2); 
//				wordUtil.addNestedTable(innerCell,innerTableData);	
				wordUtil.pageBorder();
				//wordUtil.tableContent();
				wordUtil.save("my_test_doc_" + currentDate);
				
			} catch (IOException e ) {
				e.printStackTrace();
			}
				catch (XmlException e) {
				e.printStackTrace();
			} 	
    }
}
