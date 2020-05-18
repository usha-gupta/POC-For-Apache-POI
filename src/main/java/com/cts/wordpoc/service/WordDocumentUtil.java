package com.cts.wordpoc.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageBorderOffset;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;

public class WordDocumentUtil {
	
	
	private XWPFDocument document; 
	private XWPFHeaderFooterPolicy policy; 
		
	public XWPFDocument create() throws IOException, XmlException { 
		//Create document
		this.document = new XWPFDocument(); 		
		CTSectPr sectPr = this.document.getDocument().getBody().addNewSectPr();
		this.policy = new XWPFHeaderFooterPolicy(this.document, sectPr);
		this.document.getBodyElements();
		createTOC("Table Of Contents");		
		return this.document;
	}
	
	
	public void save(String fileName) throws IOException { 
		//save document
		FileOutputStream out = new FileOutputStream(new File(fileName + ".docx"));
		this.document.write(out);
		out.close();
	}
		
		
	public void addTitle(String title, int fontSize, ParagraphAlignment alignment, String styleId) {
		//for Title
		XWPFParagraph paragraphTitle = this.document.createParagraph();
		paragraphTitle.setPageBreak(true);
		XWPFRun titleRun = paragraphTitle.createRun();
		paragraphTitle.setAlignment(alignment);	  
		titleRun.setText(title);
		titleRun.setFontSize(fontSize);
		titleRun.setBold(true);		
		paragraphTitle.setStyle(styleId);
		
	}
		
		
	public void addParagraph(String content, int fontSize) {			
		//write body content
		//titlePoint.addBreak();  //for line break
		XWPFParagraph paragraphContent = this.document.createParagraph();
		paragraphContent.setAlignment(ParagraphAlignment.BOTH);
		XWPFRun paragraphRun = paragraphContent.createRun();
		paragraphRun.setFontSize(fontSize);
		paragraphRun.setText(content);
		
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
		//table.getCTTbl().getTblPr().unsetTblBorders();
		table.setCellMargins(50, 150, 50, 150);		
		XWPFTableRow headerRow = table.getRow(0);
		XWPFTableCell headerCell;	
		
		for(int i = 0; i< headers.length; i++) {
			if (i == 0) {
				headerCell = headerRow.getCell(0);
			}
			else {
				headerCell = headerRow.addNewTableCell();
			}
			XWPFParagraph headerContent = headerCell.getParagraphs().get(0);
			headerContent.setAlignment(ParagraphAlignment.CENTER);
			headerCell.setVerticalAlignment(XWPFVertAlign.CENTER);
			XWPFRun headerContentRun = headerContent.createRun();
			headerContentRun.setColor("ffffff");
			headerContentRun.setText(headers[i]);
			headerContentRun.isBold();
			headerCell.setColor(headerColor);					
		}		
		return table;
	}
	
	
	public void addRows(XWPFTable table, String[][] tableData, String oddRowColor, String evenRowColor) {
	
		for(int i = 0; i<tableData.length; i++) {
			XWPFTableRow tableRow = table.createRow();			
			for(int j = 0; j<tableData[0].length; j++) {
				tableRow.getCell(j).setText(tableData[i][j]);
				tableRow.getCell(j).setVerticalAlignment(XWPFVertAlign.CENTER);
				if(i%2 == 0)
					tableRow.getCell(j).setColor(oddRowColor);
				else
					tableRow.getCell(j).setColor(evenRowColor);	
		    }	
		}
	}
	
	
	public void addNestedTable(XWPFTableCell cell,String[][] tableData) {		
		CTTbl  ctTbl = cell.getCTTc().addNewTbl();
		cell.getCTTc().addNewP();
		XWPFTable innerTable = new XWPFTable(ctTbl, cell);
		innerTable.getCTTbl().getTblPr().unsetTblBorders();
		innerTable.setCellMargins(50, 100, 50, 100);
		XWPFTableRow innerTableRow = null;
		XWPFTableCell innerTableColum = null;
		for(int i = 0; i < tableData.length; i++){
            if (i == 0) {
                 innerTableRow = innerTable.getRow(0);
            }
             else {
                  innerTableRow = innerTable.createRow(); 
            }
			for(int j = 0; j < tableData[0].length; j++) {			
				if ((i == 0 && j == 0) || (i != 0)) {
					innerTableColum = innerTableRow.getCell(j);
			     }
				else {
					innerTableColum = innerTableRow.addNewTableCell();		
				}
				innerTableColum.setText(tableData[i][j]);
			}
		}		
	}

	public void pageBorder() {
		CTDocument1 ctDocument = this.document.getDocument();
		CTBody ctBody = ctDocument.getBody();
		CTSectPr ctSectPr = (ctBody.isSetSectPr())?ctBody.getSectPr():ctBody.addNewSectPr();
		CTPageSz ctPageSz = (ctSectPr.isSetPgSz())?ctSectPr.getPgSz():ctSectPr.addNewPgSz();
		//paper size letter
		ctPageSz.setW(java.math.BigInteger.valueOf(Math.round(8.5 * 1440))); //8.5 inches
		ctPageSz.setH(java.math.BigInteger.valueOf(Math.round(11 * 1440))); //11 inches
		  
		//page borders
		CTPageBorders ctPageBorders = (ctSectPr.isSetPgBorders())?ctSectPr.getPgBorders():ctSectPr.addNewPgBorders();
		//ctPageBorders.setOffsetFrom(STPageBorderOffset.PAGE);
		  
		for (int b = 0; b < 4; b++) {
			CTBorder ctBorder = (ctPageBorders.isSetTop())?ctPageBorders.getTop():ctPageBorders.addNewTop();
			if (b == 1) ctBorder = (ctPageBorders.isSetBottom())?ctPageBorders.getBottom():ctPageBorders.addNewBottom();
			else if (b == 2) ctBorder = (ctPageBorders.isSetLeft())?ctPageBorders.getLeft():ctPageBorders.addNewLeft();
			else if (b == 3) ctBorder = (ctPageBorders.isSetRight())?ctPageBorders.getRight():ctPageBorders.addNewRight();
			ctBorder.setVal(STBorder.THICK);
			ctBorder.setSz(java.math.BigInteger.valueOf(10));
			ctBorder.setSpace(java.math.BigInteger.valueOf(200));
			ctBorder.setColor("000000");
		}
	}

	
	public void createTOC(String title) {
		// create a new paragraph and set title to it
	    XWPFParagraph tocPara = this.document.createParagraph();	               
	    XWPFRun TOCRun = tocPara.createRun();	       		 
	    TOCRun.setFontSize(18);
	    TOCRun.setColor("0C184C");	              
	    TOCRun.setText(title);
	    CTP ctP = tocPara.getCTP();
	    CTSimpleField toc = ctP.addNewFldSimple();
	    toc.setInstr("TOC \\h");
	    toc.setDirty(STOnOff.TRUE);
	}

	public void addCustomHeadingStyle(String styleId, int headingLevel ) {
		
		CTStyle ctStyle = CTStyle.Factory.newInstance();
		ctStyle.setStyleId(styleId);

	    CTString styleName = CTString.Factory.newInstance();
	    styleName.setVal(styleId);
	    ctStyle.setName(styleName);	
	    CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
	    indentNumber.setVal(BigInteger.valueOf(headingLevel));
	
	    // lower number > style is more prominent in the formats bar
	    ctStyle.setUiPriority(indentNumber);	
	    CTOnOff onoffnull = CTOnOff.Factory.newInstance();
	    ctStyle.setUnhideWhenUsed(onoffnull);
	
	    // style shows up in the formats bar
	    ctStyle.setQFormat(onoffnull);
	
	    // style defines a heading of the given level
	    CTPPr ppr = CTPPr.Factory.newInstance();
	    ppr.setOutlineLvl(indentNumber);
	    ctStyle.setPPr(ppr);
	
	    XWPFStyle style = new XWPFStyle(ctStyle);
	
	    // is a null op if already defined
	    XWPFStyles styles = this.document.createStyles();

	    style.setType(STStyleType.PARAGRAPH);
	    styles.addStyle(style);
	}


}