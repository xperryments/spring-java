package com.example.myweb.sampleone.hello;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

@RestController
public class HelloController {

	@RequestMapping("/writeit")
	public String sayHi() {
		try {
			XWPFDocument docx = new XWPFDocument();
			CTSectPr sectPr = docx.getDocument().getBody().addNewSectPr();
			XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);
				
			//write header content
			CTP ctpHeader = CTP.Factory.newInstance();
			CTR ctrHeader = ctpHeader.addNewR();
			CTText ctHeader = ctrHeader.addNewT();
			String headerText = "This is header";
			ctHeader.setStringValue(headerText);	
			XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, docx);
		        XWPFParagraph[] parsHeader = new XWPFParagraph[1];
		        parsHeader[0] = headerParagraph;
		        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);
		        
			//write footer content
			CTP ctpFooter = CTP.Factory.newInstance();
			CTR ctrFooter = ctpFooter.addNewR();
			CTText ctFooter = ctrFooter.addNewT();
			String footerText = "This is footer";
			ctFooter.setStringValue(footerText);	
			XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, docx);
		        XWPFParagraph[] parsFooter = new XWPFParagraph[1];
		        parsFooter[0] = footerParagraph;
			policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);
				
			//write body content
			XWPFParagraph bodyParagraph = docx.createParagraph();
			bodyParagraph.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun r = bodyParagraph.createRun();
			r.setBold(true);
			r.setText("This is body content.");
				
		        FileOutputStream out = new FileOutputStream("e:/write-test.docx");
		        docx.write(out);
		        out.close();
		        return("Done");
		    } catch (Exception ex) {
			   ex.printStackTrace();
			   return("Error");
		    }
	}
	
	@RequestMapping("/readit")
	public String readDocxFile() throws IOException, InvalidFormatException
	{
		InputStream inputStream = this.getClass().getClassLoader().getResourceAsStream("template.docx");
	    
	    XWPFDocument doc = new XWPFDocument(inputStream);
	    List<XWPFParagraph> paragraphsCell;
	    XWPFParagraph paragraphCell;
	    List<XWPFRun> runsCell;
	    XWPFRun runCell;
	    
	    //Paragraphs
	    List<XWPFParagraph> paragraphs = doc.getParagraphs();
	    XWPFParagraph paragraph = paragraphs.get(1);

	    //Runs
	    List<XWPFRun> runs = paragraph.getRuns();
	    XWPFRun run = runs.get(0);
	    run.setText(run.getText(0).replace("[Reporting Date for the Current and Prior Years]", "8/27/2017"),0);
/*	    
	    run.setFontSize(12);
	    run.setBold(true);
	    //etc.
*/
	    //Tables
	    List<XWPFTable> tables = doc.getTables();
	    XWPFTable table = tables.get(0);
	    //Rows
	    List<XWPFTableRow> rows = table.getRows();
	    XWPFTableRow row = rows.get(1);
	    //Cells
	    List<XWPFTableCell> cells = row.getTableCells();
	    XWPFTableCell cell = cells.get(0);
	    paragraphsCell = cell.getParagraphs();
	    paragraphCell = paragraphsCell.get(0);
	    runsCell = paragraphCell.getRuns();
	    runCell = runsCell.get(0);
	    runCell.setText(runCell.getText(0).replace("<Name of Client>", "Cummins"),0);
	    //cell.setParagraph(paragraphCell);
	    //etc.

	    FileOutputStream out = new FileOutputStream("e:/write-test.docx");
        doc.write(out);
        out.close();

	    return(runCell.getText(0).replace("<Name of Client>", "Cummins"));
	}
}
