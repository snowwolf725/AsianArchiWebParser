package org.snowwolf.file;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;

public class ExcelObject {
    private String anchorName;
    /**
     * Excel Stream
     */
    private InputStream inputStream;
    /**
     * POI Excel
     */
    private Excel excel;
    
    public ExcelObject(Workbook _wb){
        this.excel = new Excel(_wb);
    }
    
    public ExcelObject(InputStream inputStream){
        this.inputStream = inputStream;
        this.excel = new Excel(this.inputStream);
    }

    public ExcelObject(String anchorName , InputStream inputStream){
        this.anchorName = anchorName;
        this.inputStream = inputStream;
        this.excel = new Excel(this.inputStream);
    }
    public String getAnchorName() {
        return anchorName;
    }
    public void setAnchorName(String anchorName) {
        this.anchorName = anchorName;
    }
    public InputStream getInputStream() {
        return this.inputStream;
    }
    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }
    Excel getExcel() {
        return excel;
    }
}
