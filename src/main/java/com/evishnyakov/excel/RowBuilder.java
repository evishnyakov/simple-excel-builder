package com.evishnyakov.excel;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class RowBuilder {
    private Integer row;
    private Float rowHeightInPoints;

    public RowBuilder row(int row) {
        this.row = row;
        return this;
    }
    public RowBuilder rowHeightInPoints(Float rowHeightInPoints) {
        this.rowHeightInPoints = rowHeightInPoints;
        return this;
    }
    public void build(XSSFSheet sheet) {
        XSSFRow row = PoiUtils.getOrCreateRow(sheet, this.row);
        row.setHeightInPoints(rowHeightInPoints);
    }
}
