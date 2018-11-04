package com.evishnyakov.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

import static com.evishnyakov.excel.PoiUtils.applyBorders;

public class CellStyleCache {
    public final Map<StyleCell, CellStyle> style2cell = new HashMap<>();

    private final XSSFWorkbook workbook;
    private final FontCache fontCache;

    public CellStyleCache(XSSFWorkbook workbook, FontCache fontCache) {
        this.workbook = workbook;
        this.fontCache = fontCache;
    }

    public CellStyle getOrCreate(StyleCell style) {
        CellStyle existedCellStyle = style2cell.get(style);
        if(existedCellStyle != null) {
            return existedCellStyle;
        }
        XSSFCellStyle cellStyle = this.workbook.createCellStyle();
        if(style.getHorizontalAlignment() != null) {
            cellStyle.setAlignment(style.getHorizontalAlignment());
        }
        if(style.getFont() != null) {
            cellStyle.setFont(fontCache.getOrCreateFont(style.getFont()));
        }
        if(style.getColor() != null) {
            cellStyle.setFillForegroundColor(new XSSFColor(style.getColor(), new DefaultIndexedColorMap()));
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        if(style.getWrapText() != null) {
            cellStyle.setWrapText(style.getWrapText());
        }
        if(style.getDataFormat() != null) {
            cellStyle.setDataFormat(style.getDataFormat());
        }

        applyBorders(style, cellStyle);

        if(style.getRotation() != null) {
            cellStyle.setRotation(style.getRotation().shortValue());
        }
        if(style.getFormatPattern() != null) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(style.getFormatPattern()));
        }
        cellStyle.setVerticalAlignment(style.getVerticalAlignment() != null
                ? style.getVerticalAlignment()
                : VerticalAlignment.CENTER);
        style2cell.put(style, cellStyle);
        return cellStyle;
    }
}
