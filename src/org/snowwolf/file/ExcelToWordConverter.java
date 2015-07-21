package org.snowwolf.file;

import java.awt.Rectangle;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

import javax.xml.namespace.QName;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.VerticalText;

public class ExcelToWordConverter {

	public static void main(String[] args) {
		try {
			XWPFDocument doc = new XWPFDocument(new FileInputStream("信封-印刷品.docx"));
			List<XWPFParagraph> paragraphList = doc.getParagraphs();
			for (XWPFParagraph paragraph :paragraphList) {
				XmlObject[] textBoxObjects =  paragraph.getCTP().selectPath(
						"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' " +
				        "declare namespace wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape' .//*/wps:txbx/w:txbxContent");
				
				for (int i =0; i < textBoxObjects.length; i++) {
					QName qname = new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p");
			        XmlObject[] paraObjects = textBoxObjects[i].selectChildren(qname);

			        for (int j=0; j< paraObjects.length; j++) {
			            // Show TextBox text
			        	XWPFParagraph embeddedPara = new XWPFParagraph(CTP.Factory.parse(paraObjects[j].xmlText()), paragraph.getBody());
			            System.out.println(embeddedPara.getText());
			            
			        	// Replace text in TextBox
//			            XmlCursor cursor = paraObjects[j].newCursor();
//			            cursor.toChild(1);
//			            cursor.toChild(1);
//			            cursor.setTextValue("First");
//			            System.out.println(paraObjects[j].toString());
			        }
				}
			}
//			XmlCursor cursor = paragraphList.get(0).getCTP().newCursor();
//			XWPFParagraph new_par = doc.insertNewParagraph(cursor);
//			new_par.setPageBreak(true);
//			new_par.createRun().setText("Stupid text");
//			doc.write(new FileOutputStream("output.docx"));
			XWPFDocument doc2 = new XWPFDocument(new FileInputStream("output.docx"));
			PdfOptions options = PdfOptions.create();
			OutputStream out = new FileOutputStream(new File("output.pdf"));
	        PdfConverter.getInstance().convert(doc2, out, options);
	        convert();
		} catch (Throwable t) {
			t.printStackTrace();
		}
	}
	
	public static void convert() {
		Document document = new Document();
         try {
        	 OutputStream os = new FileOutputStream(new File("test.pdf"));
        	 XWPFDocument doc = new XWPFDocument(new FileInputStream("信封-印刷品.docx"));
        	 PdfWriter writer = PdfWriter.getInstance(document, os);
 	         //Open document
        	 document.open();
//        	 writer.setPageEmpty(true);
             document.newPage();
//             writer.setPageEmpty(true);
             List<XWPFParagraph> paragraphList = doc.getParagraphs();
             BaseFont bfChinese;
			bfChinese = BaseFont.createFont("Templates/msjhbd.ttf", BaseFont.IDENTITY_V, BaseFont.NOT_EMBEDDED);
			BaseFont bf = new Font(bfChinese , 8, Font.NORMAL).getBaseFont();
			BaseColor textColor = new BaseColor(0, 0, 0);
			PdfContentByte cb = writer.getDirectContent();
			
 			 for (XWPFParagraph paragraph :paragraphList) {
 				XmlObject[] textBoxObjects =  paragraph.getCTP().selectPath(
						"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' " +
				        "declare namespace wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape' .//*/wps:txbx/w:txbxContent");
				
				for (int i =0; i < textBoxObjects.length; i++) {
					QName qname = new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p");
			        XmlObject[] paraObjects = textBoxObjects[i].selectChildren(qname);

			        for (int j=0; j< paraObjects.length; j++) {
			            // Show TextBox text
			        	XWPFParagraph embeddedPara = new XWPFParagraph(CTP.Factory.parse(paraObjects[j].xmlText()), paragraph.getBody());
			            System.out.println(embeddedPara.getText());
			            cb.saveState();
			            cb.restoreState();
			            cb.saveState();
			            VerticalText vt = new VerticalText(writer.getDirectContent());
			            Rectangle rect = new Rectangle(10, 10, 100, 100);
	        	        vt.setVerticalLayout((float)rect.getX(), (float)rect.getY(), (float)rect.getHeight(), 12, 0);
	        	        vt.setAlignment(Element.ALIGN_CENTER);
	        	        vt.addText(new Phrase("test123", new Font(bf, 10, Font.UNDEFINED, textColor)));
	        	        vt.go();
	        	        cb.restoreState();
			        }
				}
 			 }
 			document.add(new Paragraph("中文Hello中文 World", new Font(bf, 10, Font.UNDEFINED, textColor)));
         } catch (Exception e) {  
             e.printStackTrace();  
         } finally {
        	 document.close();
         }
	}
}
