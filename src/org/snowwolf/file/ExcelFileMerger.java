package org.snowwolf.file;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.snowwolf.parser.PrintableDataItem;

public class ExcelFileMerger {
	
	private String TMP_DIR = "output";
	
	private List<PrintableDataItem> m_items = new ArrayList<PrintableDataItem>();
	
	private List<AbstractColumnProcessor> m_processors = new ArrayList<AbstractColumnProcessor>();
	
	private String m_fileName = "";
	
	private String m_templateName = "";
	
	private int m_fileCount = 0;
	
	private int m_rowCount = 0;
	
	public void setPrintInfo(String _templateName, String _fileName, int _fileCount) {
		m_templateName = _templateName;
		m_fileName = _fileName;
		m_fileCount = _fileCount;
	}
	
	public void addColumnProcesser(AbstractColumnProcessor _processor) {
		m_processors.add(_processor);
	}
	
	public void loadFiles() {
		String filename = m_fileName;
		try {
			for(int i = 1;i < m_fileCount;i++) {
				FileInputStream input_document = new FileInputStream(TMP_DIR + "/" + filename + "_" + i + ".xlsx");
				XSSFWorkbook workbook = new XSSFWorkbook(input_document);
				XSSFSheet sheet = workbook.getSheet("Sheet1");
				Row row = null;
				Cell cell = null;
				
				for(int rowIndex = 1; rowIndex <= sheet.getLastRowNum();rowIndex++) {
					row = sheet.getRow(rowIndex);
					if(row == null) {
						continue;
					} else {
						m_rowCount++;
					}
					for(int colIndex = 0; colIndex <= row.getLastCellNum();colIndex++) {
						cell = row.getCell(colIndex, Row.RETURN_NULL_AND_BLANK);
						if(cell == null) {
							continue;
						}
						PrintableDataItem item = new PrintableDataItem(m_rowCount, colIndex, cell.getStringCellValue());
						m_items.add(item);
					}
				}
				input_document.close();
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void processData() {
		for(AbstractColumnProcessor processor : m_processors) {
			processor.process(m_items);
		}
	}

	public void merge() {
		loadFiles();
		processData();
		String filename = m_fileName;
		try {
			FileInputStream input_document = new FileInputStream("Templates/" + m_templateName + ".xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(input_document);
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			Row row = null;
			Cell cell = null;
			
			for(PrintableDataItem item : m_items) {
				row = sheet.getRow(item.getRow());
				if(row == null) {
					row = sheet.createRow(item.getRow());
				}
				cell = row.getCell(item.getCol(), Row.RETURN_NULL_AND_BLANK);
				if(cell == null) {
					cell = row.createCell(item.getCol());
				}
				cell.setCellValue(item.getValue());
			}
			input_document.close();
			
			FileOutputStream output_file =new FileOutputStream(new File(TMP_DIR + "/" + filename + "_fin.xlsx"));
			workbook.write(output_file);
			output_file.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
