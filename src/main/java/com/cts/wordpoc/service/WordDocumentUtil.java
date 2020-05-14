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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

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
		XWPFTableRow tableRow0 = table.getRow(0);
		XWPFTableCell tableCell;	
		//table.getCTTbl().getTblPr().unsetTblBorders();
		int i=0;
		
		while(i< headers.length) {
			if (i == 0) {
			    tableCell = tableRow0.getCell(0);
			}
			else {
			    tableCell = tableRow0.addNewTableCell();
			}
			XWPFParagraph tableContent = tableCell.addParagraph();
			tableContent.setAlignment(ParagraphAlignment.CENTER);
			tableCell.setVerticalAlignment(XWPFVertAlign.CENTER);
			XWPFRun addContent = tableContent.createRun();
			addContent.setColor("ffffff");
			addContent.setText(headers[i]);
			addContent.isBold();
			tableCell.setColor(headerColor);
			tableCell.setVerticalAlignment(XWPFVertAlign.CENTER);			
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
			
				if(i==1 && j==2) {
					String[][] tableData1 = {{"1",  "2"}, {"3", "4"}, {"5", "6"}};
					XWPFTableCell meraCell =table.getRow(i).getCell(j);
					addNestedTable(meraCell,tableData1);	
				 }
			}	
		}
	}
	
	
	
	
	public void addNestedTable(XWPFTableCell cell,String[][] tableData) {
		
		
		XWPFTableCell tableCell= cell.getTableRow().getCell(0);
		CTTbl  ctTbl = cell.getTableRow().getCell(0).getCTTc().addNewTbl();
		ctTbl=tableCell.getCTTc().addNewTbl();
		//tableCell.removeParagraph(0);
		tableCell.getCTTc().addNewP();
		
		
		
		XWPFTable innerTable= new XWPFTable(ctTbl, tableCell);
		//innerTable.getCTTbl().getTblPr().unsetTblBorders();		
		
		
		
		XWPFTableRow innerTableRow=null;
		
		for(int i=0; i<tableData.length;i++){
			if(i!=0) {
			innerTableRow= innerTable.createRow();
			 }
			for(int j=0;j<tableData[0].length;j++) {			
				if(i==0)
				{
					if(j==0) {
						innerTableRow= innerTable.getRow(0);
						XWPFTableCell innerTableCell0=innerTableRow.getCell(0);
						innerTableCell0.setText(tableData[i][j]);
					}
					else {
						XWPFTableCell innerTableCell1=innerTableRow.addNewTableCell();
						innerTableCell1.setText(tableData[i][j]);			
					}
				}
				else {
					XWPFTableCell innerTableCell10=innerTableRow.getCell(0);
					innerTableCell10.setText(tableData[i][j]);
					XWPFTableCell innerTableCell11=innerTableRow.getCell(1);
					innerTableCell11.setText(tableData[i][j]);
					}
			}
		
		}
		
	}
	
}