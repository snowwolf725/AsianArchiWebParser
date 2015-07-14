package org.snowwolf.file;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

class Excel {

    protected Workbook wb;
    protected Sheet sheet;
    
    public Excel(Workbook _wb) {
        this.wb = _wb;
        this.sheet = wb.getSheetAt(wb.getActiveSheetIndex());
    }

    public Excel(InputStream is) {
        try {
            this.wb = WorkbookFactory.create(is);
            this.sheet = wb.getSheetAt(wb.getActiveSheetIndex());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public Sheet getSheet() {
        return sheet;
    }
    
    public Workbook getWorkbook(){
        return wb;
    }
}

