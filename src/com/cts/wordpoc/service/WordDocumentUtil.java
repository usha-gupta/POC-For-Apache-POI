package com.cts.wordpoc.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class WordDocumentUtil {
	
	
	private XWPFDocument document; 
	private XWPFHeaderFooterPolicy policy; 
		
		
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

		
	
	

}
