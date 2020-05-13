package com.cts.wordpoc.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class WordDocumentUtil {
	
	
	private XWPFDocument document; 
	private XWPFHeaderFooterPolicy policy; 
	FileOutputStream fileOutputStream = null;
		
	public XWPFDocument create() throws IOException, XmlException { 
		//Create document
		this.document = new XWPFDocument(); 		
		CTSectPr sectPr = this.document.getDocument().getBody().addNewSectPr();
		this.policy = new XWPFHeaderFooterPolicy(this.document, sectPr);
		return this.document;
	}
	
	
	public void save(String fileName) throws IOException { 
		//save document
		FileOutputStream out = new FileOutputStream(new File(fileName+".docx"));
		this.document.write(out);
		out.close();
	}
		
		
	public void addTitle(String title, int fontSize, ParagraphAlignment alignment) {
		//for Title
		XWPFParagraph paragraphTitle = this.document.createParagraph();
		XWPFRun addTitle = paragraphTitle.createRun();
		paragraphTitle.setAlignment(alignment);	  
		addTitle.setText(title);
		addTitle.setFontSize(fontSize);
		addTitle.setBold(true);		
	}
		
		
	public void addParagraph(String content, int fontSize) {
			
		//write body content
		//titlePoint.addBreak();  //for line break
		XWPFParagraph bodyParagraph = this.document.createParagraph();
		bodyParagraph.setAlignment(ParagraphAlignment.BOTH);
		XWPFRun addDescription = bodyParagraph.createRun();
		addDescription.setFontSize(fontSize);
		addDescription.setText(content);
	}
		
	public void addParagraph(String content) {
		  addParagraph(content, 12);
	      }
		
		
	public void addHeader(String header) throws IOException { 
		// here you will create header 			
		CTP ctpHeader = CTP.Factory.newInstance();
		CTR ctrHeader = ctpHeader.addNewR();
		CTText ctHeader = ctrHeader.addNewT();
		ctHeader.setStringValue(header);	
		XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, this.document);
	    XWPFParagraph[] parsHeader = new XWPFParagraph[1];
	    parsHeader[0] = headerParagraph;
	    this.policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);
	}
		

	public void addFooter(String footer) throws IOException {
		//write footer content
		CTP ctpFooter = CTP.Factory.newInstance();
		CTR ctrFooter = ctpFooter.addNewR();
		CTText ctFooter = ctrFooter.addNewT();
		ctFooter.setStringValue(footer);	
		XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, this.document);
		XWPFParagraph[] parsFooter = new XWPFParagraph[1];
		parsFooter[0] = footerParagraph;
		this.policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);
	}


	public XWPFTable addTable(String[] headers, String headerColor){
		
			// Create a Simple Table using the document.
			XWPFTable table = this.document.createTable();
			table.setCellMargins(10, 50, 100, 50);
			//table.setWidth(10);
			
			XWPFTableRow tableRow0 = table.getRow(0);
		
			XWPFTableCell tableCell0 = tableRow0.getCell(0);
			tableCell0.setColor(headerColor);
			
			
			//XWPFParagraph tableContent = this.document.createParagraph();
			XWPFParagraph tableContent = tableCell0.addParagraph();
			tableContent.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun addContent = tableContent.createRun();
			addContent.setColor("ffffff");
			addContent.isBold();
			addContent.setText(headers[0]);
			
			// Creating the First Cell
			tableCell0.setVerticalAlignment(XWPFVertAlign.CENTER);
			//tableCell0.setParagraph(tableContent);
			
			int i=1;
			
			while(i< headers.length) {
				XWPFTableCell tableCell = tableRow0.addNewTableCell();
				XWPFParagraph tableContent1 = tableCell.addParagraph();
				tableContent1.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun addContent1 = tableContent1.createRun();
				addContent1.setColor("ffffff");
				addContent1.setText(headers[i]);
				addContent1.isBold();
				
				
				
				tableCell.setColor(headerColor);
				tableCell.setVerticalAlignment(XWPFVertAlign.CENTER);
				//tableCell.setText(headers[i]);
				//tableCell.setParagraph(tableContent1);
				i++;
			}
			return table;
 
	}
	
	
	public void addRows(XWPFTable table, String[][] tableData, String oddRowColor, String evenRowColor) {
		
		for(int i=0;i<tableData.length;i++) {
			XWPFTableRow tableRow = table.createRow();
			
			for(int j=0;j<tableData[0].length;j++) {
				tableRow.getCell(j).setText(tableData[i][j]);
				tableRow.getCell(j).setVerticalAlignment(XWPFVertAlign.CENTER);
				if(i%2==0)
					tableRow.getCell(j).setColor(oddRowColor);
				else
					tableRow.getCell(j).setColor(evenRowColor);			
								
			}
			
		}

	}


}
