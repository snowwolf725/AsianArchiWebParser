package org.snowwolf.parser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReportGenerator {

	private String TMP_DIR = "output";
	
	private List<PrintableDataItem> m_item = new ArrayList<PrintableDataItem>();
	
	private String m_fileName = "";
	
	private String m_templateName = "";
	
	public void setPrintInfo(String _templateName, String _fileName, List<PrintableDataItem> _item) {
		m_templateName = _templateName;
		m_fileName = _fileName;
		m_item = _item;
	}
	
	public void genReport() {
		String filename = m_fileName;
		try {
			FileInputStream input_document = new FileInputStream("Templates/" + m_templateName + ".xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(input_document);
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			Row row = null;
			Cell cell = null;
			
			for(PrintableDataItem item : m_item) {
				row = sheet.getRow(item.getRow());
				if(row == null) {
					row = sheet.createRow(item.getRow());
				}
				cell = row.getCell(item.getCol(), Row.RETURN_NULL_AND_BLANK);
				if(cell == null) {
					cell = row.createCell(item.getCol());
					CellStyle style = workbook.createCellStyle();
					style.setAlignment(CellStyle.ALIGN_CENTER);
					cell.setCellStyle(style);
				}
				cell.setCellValue(item.getValue());
			}
			input_document.close();
			
			FileOutputStream output_file =new FileOutputStream(new File(TMP_DIR + "/" + filename + ".xlsx"));
			workbook.write(output_file);
			output_file.close();
			
//			FileInputStream fis = new FileInputStream(new File(TMP_DIR  + "/" + filename + ".xlsx"));
//			FileOutputStream fos = new FileOutputStream(new File(TMP_DIR  + "/" + filename + ".pdf"));
//			ExcelObject excel = new ExcelObject(fis);
//			Excel2Pdf pdf = new Excel2Pdf(excel , fos);
//			pdf.convert(true, false);
//			
//			String pdfPath = TMP_DIR + "/" + filename + ".pdf";
//			int ARG_COUNT = 5;
//			String [] args_1 =  new String[ARG_COUNT];
//			args_1[0]  = "-outputPrefix";
//			args_1[1]  = TMP_DIR + "/" + filename;
//			args_1[2]  = "-resolution";
//			args_1[3]  = "240";
//			args_1[4]  = pdfPath;
//			PDFToImage.main(args_1);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
