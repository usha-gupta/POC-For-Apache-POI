package com.cts.wordpoc;

import java.io.IOException;
import java.time.LocalDate;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlException;

import com.cts.wordpoc.service.DocumentBuilder;

public class App 
{
    public static void main( String[] args )
    {
    	System.out.println("hello");
        String[] tableHeaders = {"Name", "Role", "Company","Position"};
        String[][] tableData = { {"Hardware Kent", "Senior Network Architect", "NRT","Manager"}, 
        								{"Alex Adshead", "Senior Network Supporting Engineer", "","Manager"}, 
        								{"Amit Gautam",  "Releases Manager", "Cognizant ","Manager"}, 
        								{"Narendra Suraj", "project manager", "Cognizant","Manager"}};
        String[][] innerTableData = {{"Cognizant", "100"}, {"NRT", "100"},{"Cognizant", "100"}, {"NRT", "100"},{"Cognizant", "100"}};
        
        String heading1 = "Lorem ipsum is placeholder text commonly used in the graphic, print, and publishing industries for previewing layouts and visual mockups.";
        String heading2 = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.";
        String subheading1 = "Don't bother typing “lorem ipsum” into Google translate. If you already tried, you may have gotten anything from \"NATO\" to \"China\", depending on how you capitalized the letters. The bizarre translation was fodder for conspiracy theories, but Google has since updated its “lorem ipsum” translation to, boringly enough, “lorem ipsum”.";
        String subheading2 = "Now a pure snore disturbeded sum dust. He ejjnoyes, in order that somewon, also with a severe one, unless of life. May a cusstums offficer somewon nothing of a poison-filled. Until, from a twho, twho chaffinch may also pursue it, not even a lump. But as twho, as a tank; a proverb, yeast; or else they tinscribe nor. Yet yet dewlap bed. Twho may be, let him love fellows of a polecat. Now amour, the, twhose being, drunk, yet twhitch and, an enclosed valley’s always a laugh. In acquisitiendum the Furies are Earth; in (he takes up) a lump vehicles bien.”";
        String heading3 = "Lorem ipsum passages were popularized on Letraset dry-transfer sheets from the early 1970s, which were produced to be used by graphic designers for filler text.[3][4] Aldus Corporation created a version in the mid-1980s for their desktop publishing program PageMaker.[4]";
        String subheading3 = "First, the source of Lorem Ipsum—tracked down by Hampden-Sydney Director of Publications Richard McClintock---is Roman lawyer, statesmen, and philosopher Cicero, from an essay called “On the Extremes of Good and Evil,” or De Finibus Bonorum et Malorum.";
        
        String headingId="Heading1";
        String subHeadingId="subHeading";
        String level2Heading="level2Heading";
        String level3Heading="level3Heading";
        DocumentBuilder builder = new DocumentBuilder(); 
		
			try {
				XWPFDocument doc = builder.create();
				
				builder.addCustomHeadingStyle(headingId, 0);
				builder.addCustomHeadingStyle(subHeadingId, 1);
				builder.addCustomHeadingStyle(level2Heading, 2);
				builder.addCustomHeadingStyle(level3Heading, 3);
				
				
				builder.addTitle("Lorem Ipsum", 16, ParagraphAlignment.LEFT,headingId ); 
				builder.addParagraph(heading1); 


				builder.addTitle("LOREM IPSUM GENERATOR", 16, ParagraphAlignment.LEFT, headingId);
				builder.addParagraph(heading2, 14);		

				builder.addTitle("INTERPRETING NONSENSE", 16, ParagraphAlignment.LEFT, subHeadingId);
				builder.addParagraph(subheading1, 14);	

				builder.addTitle("Boparai's version:", 16, ParagraphAlignment.LEFT, subHeadingId);
				builder.addParagraph(subheading2, 14);
				
				builder.addTitle("GENERATOR:", 14, ParagraphAlignment.LEFT, level2Heading);
				builder.addParagraph(subheading2, 14);
				
				builder.addTitle("Variations", 16, ParagraphAlignment.LEFT, headingId);
				builder.addParagraph(heading3, 14);		

				builder.addTitle(" invented Lorem Ipsum", 16, ParagraphAlignment.LEFT, subHeadingId);
				builder.addParagraph(subheading3, 14);				
				
//				wordUtil.addParagraph("This is my default size paragraph"); 
//				wordUtil.addParagraph("This is my large font size paragraph", 14);				
				LocalDate currentDate = LocalDate.now();
//				wordUtil.addHeader("This is my header....");
//				wordUtil.addFooter("This is my footer");				
				XWPFTable table = builder.addTable(tableHeaders, "4d82be",DocumentBuilder.AUTO_FIT_WINDOW);
				//XWPFTable table = wordUtil.addTable(tableHeaders, "4d82be");
				builder.addRows(table, tableData, "dbe5f1", "feffff");
//				XWPFTableCell innerCell = table.getRow(2).getCell(2); 
//				wordUtil.addNestedTable(innerCell,innerTableData);	
				builder.pageBorder();
				builder.save("my_test_doc_" + currentDate);
				
			} catch (IOException e ) {
				e.printStackTrace();
			}
				catch (XmlException e) {
				e.printStackTrace();
			} 	
    }
}
