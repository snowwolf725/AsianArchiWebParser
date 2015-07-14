package org.snowwolf.file;

import java.awt.Dimension;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Sheet;

public class POIImage {
    protected Dimension dimension;
    protected byte[] bytes;
    protected ClientAnchor anchor;
    
    public POIImage getCellImage(Cell cell) {
        Sheet sheet = cell.getSheet();
//      Workbook wb = sheet.getWorkbook();
//      List<PictureData> pictures = (List<PictureData>) wb.getAllPictures();
        if (sheet instanceof HSSFSheet) {
            HSSFSheet hssfSheet = (HSSFSheet) sheet;
            List<HSSFShape> shapes = hssfSheet.getDrawingPatriarch().getChildren();
            for (HSSFShape shape : shapes) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    PictureData data = pic.getPictureData();
                    int row1 = anchor.getRow1();
                    int col1 = anchor.getCol1();
                    if(row1 == cell.getRowIndex() && col1 == cell.getColumnIndex()){
                        dimension = pic.getImageDimension();
                        this.anchor = anchor;
                        this.bytes = data.getData();
                    }
                }
            }
        }
        return this;
    }
    
    public Dimension getDimension() {
        return dimension;
    }

    public void setDimension(Dimension dimension) {
        this.dimension = dimension;
    }

    public byte[] getBytes() {
        return bytes;
    }
    
    public void setBytes(byte[] bytes) {
        this.bytes = bytes;
    }
    
    public ClientAnchor getAnchor() {
        return anchor;
    }

    public void setAnchor(ClientAnchor anchor) {
        this.anchor = anchor;
    }
}
